using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using ContentType = OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.IO;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectContentType : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Content Types"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // if this is a sub site then we're not provisioning content types. Technically this can be done but it's not a recommended practice
                if (web.IsSubSite())
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Context_web_is_subweb__Skipping_content_types_);
                    return parser;
                }

                web.Context.Load(web.ContentTypes, ct => ct.IncludeWithDefaultProperties(c => c.StringId, c => c.FieldLinks,
                                                                                         c => c.FieldLinks.Include(fl => fl.Id, fl => fl.Required, fl => fl.Hidden)));
                web.Context.Load(web.Fields, fld => fld.IncludeWithDefaultProperties(f => f.Id));

                web.Context.ExecuteQueryRetry();

                var existingCTs = web.ContentTypes.ToList();
                var existingFields = web.Fields.ToList();

                foreach (var ct in template.ContentTypes.OrderBy(ct => ct.Id)) // ordering to handle references to parent content types that can be in the same template
                {
                    var existingCT = existingCTs.FirstOrDefault(c => c.StringId.Equals(ct.Id, StringComparison.OrdinalIgnoreCase));
                    if (existingCT == null)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Creating_new_Content_Type___0_____1_, ct.Id, ct.Name);
                        var newCT = CreateContentType(web, ct, parser, template.Connector ?? null, existingCTs, existingFields);
                        if (newCT != null)
                        {
                            existingCTs.Add(newCT);
                        }
                    }
                    else
                    {
                        if (ct.Overwrite)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Recreating_existing_Content_Type___0_____1_, ct.Id, ct.Name);

                            existingCT.DeleteObject();
                            web.Context.ExecuteQueryRetry();
                            var newCT = CreateContentType(web, ct, parser, template.Connector ?? null, existingCTs, existingFields);
                            if (newCT != null)
                            {
                                existingCTs.Add(newCT);
                            }
                        }
                        else
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Updating_existing_Content_Type___0_____1_, ct.Id, ct.Name);
                            UpdateContentType(web, existingCT, ct, parser, scope);
                        }
                    }
                }
            }
            return parser;
        }


        private static void UpdateContentType(Web web, Microsoft.SharePoint.Client.ContentType existingContentType, ContentType templateContentType, TokenParser parser, PnPMonitoredScope scope)
        {
            var isDirty = false;
            if (existingContentType.Hidden != templateContentType.Hidden)
            {
                scope.LogPropertyUpdate("Hidden");
                existingContentType.Hidden = templateContentType.Hidden;
                isDirty = true;
            }
            if (existingContentType.ReadOnly != templateContentType.ReadOnly)
            {
                scope.LogPropertyUpdate("ReadOnly");
                existingContentType.ReadOnly = templateContentType.ReadOnly;
                isDirty = true;
            }
            if (existingContentType.Sealed != templateContentType.Sealed)
            {
                scope.LogPropertyUpdate("Sealed");
                existingContentType.Sealed = templateContentType.Sealed;
                isDirty = true;
            }
            if (templateContentType.Description != null && existingContentType.Description != parser.ParseString(templateContentType.Description))
            {
                scope.LogPropertyUpdate("Description");
                existingContentType.Description = parser.ParseString(templateContentType.Description);
                isDirty = true;
            }
            if (templateContentType.DocumentTemplate != null && existingContentType.DocumentTemplate != parser.ParseString(templateContentType.DocumentTemplate))
            {
                scope.LogPropertyUpdate("DocumentTemplate");
                existingContentType.DocumentTemplate = parser.ParseString(templateContentType.DocumentTemplate);
                isDirty = true;
            }
            if (existingContentType.Name != parser.ParseString(templateContentType.Name))
            {
                scope.LogPropertyUpdate("Name");
                existingContentType.Name = parser.ParseString(templateContentType.Name);
                isDirty = true;
                // CT is being renamed, add an extra token to the tokenparser
                parser.AddToken(new ContentTypeIdToken(web, existingContentType.Name, existingContentType.StringId));
            }
            if (templateContentType.Group != null && existingContentType.Group != parser.ParseString(templateContentType.Group))
            {
                scope.LogPropertyUpdate("Group");
                existingContentType.Group = parser.ParseString(templateContentType.Group);
                isDirty = true;
            }
            if (templateContentType.DisplayFormUrl != null && existingContentType.DisplayFormUrl != parser.ParseString(templateContentType.DisplayFormUrl))
            {
                scope.LogPropertyUpdate("DisplayFormUrl");
                existingContentType.DisplayFormUrl = parser.ParseString(templateContentType.DisplayFormUrl);
                isDirty = true;
            }
            if (templateContentType.EditFormUrl != null && existingContentType.EditFormUrl != parser.ParseString(templateContentType.EditFormUrl))
            {
                scope.LogPropertyUpdate("EditFormUrl");
                existingContentType.EditFormUrl = parser.ParseString(templateContentType.EditFormUrl);
                isDirty = true;
            }
            if (templateContentType.NewFormUrl != null && existingContentType.NewFormUrl != parser.ParseString(templateContentType.NewFormUrl))
            {
                scope.LogPropertyUpdate("NewFormUrl");
                existingContentType.NewFormUrl = parser.ParseString(templateContentType.NewFormUrl);
                isDirty = true;
            }
#if !CLIENTSDKV15
            if (templateContentType.Name.ContainsResourceToken())
            {
                existingContentType.NameResource.SetUserResourceValue(templateContentType.Name, parser);
                isDirty = true;
            }
            if (templateContentType.Description.ContainsResourceToken())
            {
                existingContentType.DescriptionResource.SetUserResourceValue(templateContentType.Description, parser);
                isDirty = true;
            }
#endif
            if (isDirty)
            {
                existingContentType.Update(true);
                web.Context.ExecuteQueryRetry();
            }
            // Delta handling
            existingContentType.EnsureProperty(c => c.FieldLinks);
            List<Guid> targetIds = existingContentType.FieldLinks.AsEnumerable().Select(c1 => c1.Id).ToList();
            List<Guid> sourceIds = templateContentType.FieldRefs.Select(c1 => c1.Id).ToList();

            var fieldsNotPresentInTarget = sourceIds.Except(targetIds).ToArray();

            if (fieldsNotPresentInTarget.Any())
            {
                foreach (var fieldId in fieldsNotPresentInTarget)
                {
                    var fieldRef = templateContentType.FieldRefs.Find(fr => fr.Id == fieldId);
                    var field = web.Fields.GetById(fieldId);
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Adding_field__0__to_content_type, fieldId);
                    web.AddFieldToContentType(existingContentType, field, fieldRef.Required, fieldRef.Hidden);
                }
            }

            isDirty = false;
            foreach (var fieldId in targetIds.Intersect(sourceIds))
            {
                var fieldLink = existingContentType.FieldLinks.FirstOrDefault(fl => fl.Id == fieldId);
                var fieldRef = templateContentType.FieldRefs.Find(fr => fr.Id == fieldId);
                if (fieldRef != null)
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Field__0__exists_in_content_type, fieldId);
                    if (fieldLink.Required != fieldRef.Required)
                    {
                        scope.LogPropertyUpdate("Required");
                        fieldLink.Required = fieldRef.Required;
                        isDirty = true;
                    }
                    if (fieldLink.Hidden != fieldRef.Hidden)
                    {
                        scope.LogPropertyUpdate("Hidden");
                        fieldLink.Hidden = fieldRef.Hidden;
                        isDirty = true;
                    }
                }
            }

            // The new CT is a DocumentSet, and the target should be, as well
            if (templateContentType.DocumentSetTemplate != null)
            {
                if (!Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.IsChildOfDocumentSetContentType(web.Context, existingContentType).Value)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_ContentTypes_InvalidDocumentSet_Update_Request, existingContentType.Id, existingContentType.Name);
                }
                else
                {
                    Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate templateToUpdate =
                        Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.GetDocumentSetTemplate(web.Context, existingContentType);

                    // TODO: Implement Delta Handling
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ContentTypes_DocumentSet_DeltaHandling_OnHold, existingContentType.Id, existingContentType.Name);
                }
            }

            if (isDirty)
            {
                existingContentType.Update(true);
                web.Context.ExecuteQueryRetry();
            }
        }

        private static Microsoft.SharePoint.Client.ContentType CreateContentType(Web web, ContentType templateContentType, TokenParser parser, FileConnectorBase connector,
            List<Microsoft.SharePoint.Client.ContentType> existingCTs = null, List<Microsoft.SharePoint.Client.Field> existingFields = null)
        {
            var name = parser.ParseString(templateContentType.Name);
            var description = parser.ParseString(templateContentType.Description);
            var id = parser.ParseString(templateContentType.Id);
            var group = parser.ParseString(templateContentType.Group);

            var createdCT = web.CreateContentType(name, description, id, group);
            foreach (var fieldRef in templateContentType.FieldRefs)
            {
                var field = web.Fields.GetById(fieldRef.Id);
                web.AddFieldToContentType(createdCT, field, fieldRef.Required, fieldRef.Hidden);
            }

            // Add new CTs
            parser.AddToken(new ContentTypeIdToken(web, name, id));

#if !CLIENTSDKV15
            // Set resources
            if (templateContentType.Name.ContainsResourceToken())
            {
                createdCT.NameResource.SetUserResourceValue(templateContentType.Name, parser);
            }
            if(templateContentType.Description.ContainsResourceToken())
            {
                createdCT.DescriptionResource.SetUserResourceValue(templateContentType.Description, parser);
            }
#endif
            //Reorder the elements so that the new created Content Type has the same order as defined in the
            //template. The order can be different if the new Content Type inherits from another Content Type.
            //In this case the new Content Type has all field of the original Content Type and missing fields 
            //will be added at the end. To fix this issue we ordering the fields once more.
            createdCT.FieldLinks.Reorder(templateContentType.FieldRefs.Select(fld => fld.Name).ToArray());
            createdCT.Update(true);
            web.Context.ExecuteQueryRetry();

            createdCT.ReadOnly = templateContentType.ReadOnly;
            createdCT.Hidden = templateContentType.Hidden;
            createdCT.Sealed = templateContentType.Sealed;
            if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.DocumentTemplate)))
            {
                createdCT.DocumentTemplate = parser.ParseString(templateContentType.DocumentTemplate);
            }
            if (!String.IsNullOrEmpty(templateContentType.NewFormUrl))
            {
                createdCT.NewFormUrl = templateContentType.NewFormUrl;
            }
            if (!String.IsNullOrEmpty(templateContentType.EditFormUrl))
            {
                createdCT.EditFormUrl = templateContentType.EditFormUrl;
            }
            if (!String.IsNullOrEmpty(templateContentType.DisplayFormUrl))
            {
                createdCT.DisplayFormUrl = templateContentType.DisplayFormUrl;
            }

            // If the CT is a DocumentSet
            if (templateContentType.DocumentSetTemplate != null)
            {
                // Retrieve a reference to the DocumentSet Content Type
                Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate documentSetTemplate =
                    Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.GetDocumentSetTemplate(web.Context, createdCT);

                if (!String.IsNullOrEmpty(templateContentType.DocumentSetTemplate.WelcomePage))
                {
                    // TODO: Customize the WelcomePage of the DocumentSet
                }

                foreach (String ctId in templateContentType.DocumentSetTemplate.AllowedContentTypes)
                {
                    Microsoft.SharePoint.Client.ContentType ct = existingCTs.FirstOrDefault(c => c.StringId == ctId);
                    if (ct != null)
                    {
                        documentSetTemplate.AllowedContentTypes.Add(ct.Id);
                    }
                }

                foreach (var doc in templateContentType.DocumentSetTemplate.DefaultDocuments)
                {
                    Microsoft.SharePoint.Client.ContentType ct = existingCTs.FirstOrDefault(c => c.StringId == doc.ContentTypeId);
                    if (ct != null)
                    {
                        using (Stream fileStream = connector.GetFileStream(doc.FileSourcePath))
                        {
                            documentSetTemplate.DefaultDocuments.Add(doc.Name, ct.Id, ReadFullStream(fileStream));
                        }
                    }
                }

                foreach (var sharedField in templateContentType.DocumentSetTemplate.SharedFields)
                {
                    Microsoft.SharePoint.Client.Field field = existingFields.FirstOrDefault(f => f.Id == sharedField);
                    if (field != null)
                    {
                        documentSetTemplate.SharedFields.Add(field);
                    }
                }

                foreach (var welcomePageField in templateContentType.DocumentSetTemplate.WelcomePageFields)
                {
                    Microsoft.SharePoint.Client.Field field = existingFields.FirstOrDefault(f => f.Id == welcomePageField);
                    if (field != null)
                    {
                        documentSetTemplate.WelcomePageFields.Add(field);
                    }
                }

                documentSetTemplate.Update(true);
                web.Context.ExecuteQueryRetry();
            }

            web.Context.Load(createdCT);
            web.Context.ExecuteQueryRetry();

            return createdCT;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // if this is a sub site then we're not creating content type entities. 
                if (web.IsSubSite())
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Context_web_is_subweb__Skipping_content_types_);
                    return template;
                }

                template.ContentTypes.AddRange(GetEntities(web, scope));

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate, scope);
                }
            }
            return template;
        }

        private IEnumerable<ContentType> GetEntities(Web web, PnPMonitoredScope scope)
        {
            var cts = web.ContentTypes;
            web.Context.Load(cts, ctCollection => ctCollection.IncludeWithDefaultProperties(ct => ct.FieldLinks));
            web.Context.ExecuteQueryRetry();

            List<ContentType> ctsToReturn = new List<ContentType>();

            foreach (var ct in cts)
            {
                if (!BuiltInContentTypeId.Contains(ct.StringId))
                {
                    ContentType newCT = new ContentType
                        (ct.StringId,
                        ct.Name,
                        ct.Description,
                        ct.Group,
                        ct.Sealed,
                        ct.Hidden,
                        ct.ReadOnly,
                        ct.DocumentTemplate,
                        false,
                            (from fieldLink in ct.FieldLinks.AsEnumerable<FieldLink>()
                             select new FieldRef(fieldLink.Name)
                             {
                                 Id = fieldLink.Id,
                                 Hidden = fieldLink.Hidden,
                                 Required = fieldLink.Required,
                             })
                        )
                    {
                        DisplayFormUrl = ct.DisplayFormUrl,
                        EditFormUrl = ct.EditFormUrl,
                        NewFormUrl = ct.NewFormUrl,
                    };

                    // If the Content Type is a DocumentSet
                    if (Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.IsChildOfDocumentSetContentType(web.Context, ct).Value ||
                        ct.StringId.StartsWith(BuiltInContentTypeId.DocumentSet)) // TODO: This is kind of an hack... we should find a better solution ...
                    {
                        Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate documentSetTemplate =
                            Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.GetDocumentSetTemplate(web.Context, ct);

                        // Retrieve the Document Set
                        web.Context.Load(documentSetTemplate,
                            t => t.AllowedContentTypes,
                            t => t.DefaultDocuments,
                            t => t.SharedFields,
                            t => t.WelcomePageFields);
                        web.Context.ExecuteQueryRetry();

                        newCT.DocumentSetTemplate = new DocumentSetTemplate(
                            null, // TODO: WelcomePage not yet supported
                            (from allowedCT in documentSetTemplate.AllowedContentTypes.AsEnumerable()
                             select allowedCT.StringValue).ToArray(),
                            (from defaultDocument in documentSetTemplate.DefaultDocuments.AsEnumerable()
                             select new DefaultDocument
                             {
                                 ContentTypeId = defaultDocument.ContentTypeId.StringValue,
                                 Name = defaultDocument.Name,
                                 FileSourcePath = String.Empty, // TODO: How can we extract the proper file?!
                             }).ToArray(),
                            (from sharedField in documentSetTemplate.SharedFields.AsEnumerable()
                             select sharedField.Id).ToArray(),
                            (from welcomePageField in documentSetTemplate.WelcomePageFields.AsEnumerable()
                             select welcomePageField.Id).ToArray()
                        );
                    }

                    ctsToReturn.Add(newCT);
                }
            }
            return ctsToReturn;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate, PnPMonitoredScope scope)
        {
            foreach (var ct in baseTemplate.ContentTypes)
            {
                var index = template.ContentTypes.FindIndex(f => f.Id.Equals(ct.Id, StringComparison.OrdinalIgnoreCase));
                if (index > -1)
                {
                    template.ContentTypes.RemoveAt(index);
                }
                else
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Adding_content_type_to_template___0_____1_, ct.Id, ct.Name);
                }

            }

            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.ContentTypes.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }

        private static Byte[] ReadFullStream(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream mem = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    mem.Write(buffer, 0, read);
                }
                return mem.ToArray();
            }
        }
    }
}
