using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using ContentType = OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType;
using Field = OfficeDevPnP.Core.Framework.Provisioning.Model.Field;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectContentType : ObjectHandlerBase
    {
        private FieldAndListProvisioningStepHelper.Step _step;

        public override string Name
        {
#if DEBUG
            get { return $"Content Types ({_step})"; }
#else
			get { return $"Content Types"; }
#endif
        }

        public override string InternalName => "ContentTypes";

        public ObjectContentType(FieldAndListProvisioningStepHelper.Step step)
        {
            this._step = step;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Check if this is not a noscript site as we're not allowed to update some properties
                bool isNoScriptSite = web.IsNoScriptSite();

                web.Context.Load(web.ContentTypes, ct => ct.IncludeWithDefaultProperties(c => c.StringId, c => c.FieldLinks,
                                                                                         c => c.FieldLinks.Include(fl => fl.Id, fl => fl.Required, fl => fl.Hidden)));
                web.Context.Load(web.Fields, fld => fld.IncludeWithDefaultProperties(f => f.Id, f => f.SchemaXml));

                web.Context.ExecuteQueryRetry();

                var existingCTs = web.ContentTypes.ToList();
                var existingFields = web.Fields.ToList();
                var currentCtIndex = 0;

                var doProvision = true;
                if (web.IsSubSite() && !applyingInformation.ProvisionContentTypesToSubWebs)
                {
                    WriteMessage("This template contains content types and you are provisioning to a subweb. If you still want to provision these content types, set the ProvisionContentTypesToSubWebs property to true.", ProvisioningMessageType.Warning);
                    doProvision = false;
                }
                if (doProvision)
                {
                    foreach (var ct in template.ContentTypes.OrderBy(ct => ct.Id)) // ordering to handle references to parent content types that can be in the same template
                    {
                        currentCtIndex++;

                        WriteSubProgress("Content Type", ct.Name, currentCtIndex, template.ContentTypes.Count);
                        var existingCT = existingCTs.FirstOrDefault(c => c.StringId.Equals(ct.Id, StringComparison.OrdinalIgnoreCase));
                        if (existingCT == null)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Creating_new_Content_Type___0_____1_, ct.Id, ct.Name);
                            var newCT = CreateContentType(web, template, ct, parser, template.Connector, scope, existingCTs, existingFields, isNoScriptSite);
                            if (newCT != null)
                            {
                                existingCTs.Add(newCT);
                                existingCT = newCT;
                            }
                        }
                        else
                        {
                            if (ct.Overwrite && this._step == FieldAndListProvisioningStepHelper.Step.ListAndStandardFields)
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Recreating_existing_Content_Type___0_____1_, ct.Id, ct.Name);

                                existingCT.DeleteObject();
                                web.Context.ExecuteQueryRetry();
                                var newCT = CreateContentType(web, template, ct, parser, template.Connector, scope, existingCTs, existingFields, isNoScriptSite);
                                if (newCT != null)
                                {
                                    existingCTs.Add(newCT);
                                    existingCT = newCT;
                                }
                            }
                            else
                            {
                                // We can't update a sealed or read only content type unless we change the value to false
                                if ((!existingCT.Sealed || !ct.Sealed) && (!existingCT.ReadOnly || !ct.ReadOnly))
                                {
                                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Updating_existing_Content_Type___0_____1_, ct.Id, ct.Name);
                                    UpdateContentType(web, template, existingCT, ct, parser, template.Connector, scope, existingCTs, existingFields, isNoScriptSite);
                                }
                                else
                                {
                                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Updating_existing_Content_Type_SealedOrReadOnly, ct.Id, ct.Name);
                                }
                            }
                        }

                        // Set ReadOnly as the last thing because a ReadOnly content type cannot be updated
                        if (this._step == FieldAndListProvisioningStepHelper.Step.LookupFields && existingCT.ReadOnly == false && ct.ReadOnly == true)
                        {
                            scope.LogPropertyUpdate("ReadOnly");
                            existingCT.ReadOnly = ct.ReadOnly;

                            existingCT.Update(false);
                            existingCT.Context.ExecuteQueryRetry();
                        }
                    }
                }
            }
            WriteMessage($"Done processing Content Types", ProvisioningMessageType.Completed);
            return parser;
        }

        private void UpdateContentType(
            Web web,
            ProvisioningTemplate template,
            Microsoft.SharePoint.Client.ContentType existingContentType,
            ContentType templateContentType,
            TokenParser parser,
            FileConnectorBase connector,
            PnPMonitoredScope scope,
            List<Microsoft.SharePoint.Client.ContentType> existingCTs = null,
            List<Microsoft.SharePoint.Client.Field> existingFields = null,
            bool isNoScriptSite = false
            )
        {
            var isDirty = false;
            var reOrderFields = false;
            var name = parser.ParseString(templateContentType.Name);

            if (existingContentType.Hidden != templateContentType.Hidden)
            {
                scope.LogPropertyUpdate("Hidden");
                existingContentType.Hidden = templateContentType.Hidden;
                isDirty = true;
            }
            // Only change ReadOnly here, if change is from True => False (not ReadOnly)
            // If change is ReadOnly = True, it will be set later
            if (existingContentType.ReadOnly == true && templateContentType.ReadOnly == false)
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
                var oldName = existingContentType.Name;
                scope.LogPropertyUpdate("Name");
                existingContentType.Name = parser.ParseString(templateContentType.Name);
                isDirty = true;
                // CT is being renamed, add an extra token to the tokenparser
                parser.RemoveToken(new ContentTypeIdToken(web, oldName, existingContentType.StringId));
                parser.AddToken(new ContentTypeIdToken(web, existingContentType.Name, existingContentType.StringId));
            }
            if (templateContentType.Group != null && existingContentType.Group != parser.ParseString(templateContentType.Group))
            {
                scope.LogPropertyUpdate("Group");
                existingContentType.Group = parser.ParseString(templateContentType.Group);
                isDirty = true;
            }

            if (!isNoScriptSite)
            {
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
            }
            else
            {
                if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.DisplayFormUrl)) ||
                    !string.IsNullOrEmpty(parser.ParseString(templateContentType.EditFormUrl)) ||
                    !string.IsNullOrEmpty(parser.ParseString(templateContentType.NewFormUrl)))
                {
                    // log message
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ContentTypes_SkipCustomFormUrls, existingContentType.Name);
                }
            }

#if !SP2013
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
                // Default to false as there is no reason to update children on CT property changes.
                existingContentType.Update(false);
                web.Context.ExecuteQueryRetry();
            }

            // Set flag to reorder fields CT fields are not equal to template fields
            var existingFieldNames = existingContentType.FieldLinks.AsEnumerable().Select(fld => fld.Name).ToArray();
            var ctFieldNames = templateContentType.FieldRefs.Select(fld => parser.ParseString(fld.Name)).ToArray();
            reOrderFields = ctFieldNames.Length > 0 && !existingFieldNames.SequenceEqual(ctFieldNames);

            // Delta handling
            existingContentType.EnsureProperty(c => c.FieldLinks);
            var targetIds = existingContentType.FieldLinks.AsEnumerable().Select(c1 => c1.Id).ToList();
            var sourceIds = templateContentType.FieldRefs.Select(c1 => c1.Id).ToList();

            var fieldsNotPresentInTarget = sourceIds.Except(targetIds).ToArray();

            // Should child content types be updated.
            bool UpdateChildren()
            {
                if (fieldsNotPresentInTarget.Any())
                {
                    return !templateContentType.FieldRefs.All(f => f.UpdateChildren == false);
                }

                return true;
            }

            if (fieldsNotPresentInTarget.Any())
            {
                // Set flag to reorder fields when new fields are added.
                reOrderFields = true;

                foreach (var fieldId in fieldsNotPresentInTarget)
                {
                    var fieldRef = templateContentType.FieldRefs.First(fr => fr.Id == fieldId);

                    var templateField = template.SiteFields.FirstOrDefault(tf => (Guid)XElement.Parse(parser.ParseString(tf.SchemaXml)).Attribute("ID") == fieldRef.Id);
                    var fieldStep = templateField != null ? templateField.GetFieldProvisioningStep(parser) : FieldAndListProvisioningStepHelper.Step.ListAndStandardFields;
                    if (fieldStep != _step) continue; // Do not handle this field at this step

                    Microsoft.SharePoint.Client.Field field = null;
                    if (_step == FieldAndListProvisioningStepHelper.Step.LookupFields
                        && templateField != null
                        && XElement.Parse(parser.ParseString(templateField.SchemaXml)).Attribute("FieldRef") != null)
                    {
                        // Because the id of dependent lookup cannot be set and is autogenerated,
                        // we have to retrieve the actual field id and convert it into a token
                        var mappedFieldId = Guid.Parse(parser.ParseString(fieldRef.Id.ToString("D")));
                        field = web.AvailableFields.GetById(mappedFieldId);
                    }
                    else
                    {
                        field = web.AvailableFields.GetById(fieldRef.Id);
                    }

                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Adding_field__0__to_content_type, fieldId);
                    web.AddFieldToContentType(existingContentType, field, 
                        fieldRef.Required, 
                        fieldRef.Hidden, 
                        fieldRef.UpdateChildren
#if !SP2013 && !SP2016
                        ,
                        fieldRef.ShowInDisplayForm,
                        fieldRef.ReadOnly
#endif
                        );
                }
            }

            // Reorder fields
            if (reOrderFields)
            {
                existingContentType.FieldLinks.Reorder(ctFieldNames);
                isDirty = true;
            }

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
                var isChildOfDocumentSetContentType = Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.IsChildOfDocumentSetContentType(web.Context, existingContentType);
                web.Context.ExecuteQueryRetry();

                if (!isChildOfDocumentSetContentType.Value)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_ContentTypes_InvalidDocumentSet_Update_Request, existingContentType.Id, existingContentType.Name);
                }
                else
                {
                    // Retrieve a reference to the DocumentSet Content Type
                    Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate documentSetTemplate =
                        Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.GetDocumentSetTemplate(web.Context, existingContentType);

                    // Keep a flag if changes have been made to the document set of the content type
                    var documentSetIsDirty = false;

                    // Load the collections to allow for deletion scenarions
                    web.Context.Load(documentSetTemplate, d => d.AllowedContentTypes, d => d.DefaultDocuments, d => d.SharedFields, d => d.WelcomePageFields);
                    web.Context.ExecuteQueryRetry();

                    if (!String.IsNullOrEmpty(templateContentType.DocumentSetTemplate.WelcomePage))
                    {
                        // TODO: Customize the WelcomePage of the DocumentSet
                    }

                    // AllowedContentTypes
                    // Add additional content types to the set of allowed content types
                    foreach (string ctId in templateContentType.DocumentSetTemplate.AllowedContentTypes)
                    {
                        // Validate if the content type is not part of the document set content types yet
                        if (documentSetTemplate.AllowedContentTypes.All(d => d.StringValue != ctId))
                        {
                            Microsoft.SharePoint.Client.ContentType ct = existingCTs.FirstOrDefault(c => c.StringId == ctId);
                            if (ct != null)
                            {
                                documentSetTemplate.AllowedContentTypes.Add(ct.Id);
                                documentSetIsDirty = true;
                            }
                        }
                    }

                    // DefaultDocuments
                    if (!isNoScriptSite)
                    {
                        foreach (var doc in templateContentType.DocumentSetTemplate.DefaultDocuments)
                        {                                
                            // Ensure the default document is not part of the document set yet
                            if (documentSetTemplate.DefaultDocuments.All(d => d.Name != doc.Name))
                            {
                                Microsoft.SharePoint.Client.ContentType ct = existingCTs.FirstOrDefault(c => c.StringId == doc.ContentTypeId);
                                if (ct != null)
                                {
                                    using (Stream fileStream = connector.GetFileStream(doc.FileSourcePath))
                                    {
                                        documentSetTemplate.DefaultDocuments.Add(doc.Name, ct.Id, ReadFullStream(fileStream));
                                        documentSetIsDirty = true;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (templateContentType.DocumentSetTemplate.DefaultDocuments.Any())
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ContentTypes_SkipDocumentSetDefaultDocuments, name);
                        }
                    }

                    // SharedFields
                    foreach (var sharedField in templateContentType.DocumentSetTemplate.SharedFields)
                    {                            
                        // Ensure the shared field is not part of the document set yet
                        if (documentSetTemplate.SharedFields.All(f => f.Id != sharedField))
                        {
                            Microsoft.SharePoint.Client.Field field = existingFields.FirstOrDefault(f => f.Id == sharedField);
                            if (field != null)
                            {
                                documentSetTemplate.SharedFields.Add(field);
                                documentSetIsDirty = true;
                            }
                        }
                    }

                    // WelcomePageFields
                    foreach (var welcomePageField in templateContentType.DocumentSetTemplate.WelcomePageFields)
                    {
                        // Ensure the welcomepage field is not part of the document set yet
                        if (documentSetTemplate.WelcomePageFields.All(w => w.Id != welcomePageField))
                        {
                            Microsoft.SharePoint.Client.Field field = existingFields.FirstOrDefault(f => f.Id == welcomePageField);
                            if (field != null)
                            {
                                documentSetTemplate.WelcomePageFields.Add(field);
                                documentSetIsDirty = true;
                            }
                        }
                    }

                    if (documentSetIsDirty)
                    {
                        documentSetTemplate.Update(true);
                        isDirty = true;
                    }
                }
            }

            if (isDirty)
            {
                scope.LogDebug("Update child Content Types: {0}", UpdateChildren());
                existingContentType.Update(UpdateChildren());
                web.Context.ExecuteQueryRetry();
            }
        }

        private Microsoft.SharePoint.Client.ContentType CreateContentType(
            Web web,
            ProvisioningTemplate template,
            ContentType templateContentType,
            TokenParser parser,
            FileConnectorBase connector,
            PnPMonitoredScope scope,
            List<Microsoft.SharePoint.Client.ContentType> existingCTs = null,
            List<Microsoft.SharePoint.Client.Field> existingFields = null,
            bool isNoScriptSite = false)
        {
            var name = parser.ParseString(templateContentType.Name);
            var description = parser.ParseString(templateContentType.Description);
            var id = parser.ParseString(templateContentType.Id);
            var group = parser.ParseString(templateContentType.Group);

            var createdCT = web.CreateContentType(name, description, id, group);
            createdCT.EnsureProperties(ct => ct.ReadOnly, ct => ct.Hidden, ct => ct.Sealed);

            List<FieldRef> fieldsRefsToProcess = new List<FieldRef>();
            foreach (FieldRef fr in templateContentType.FieldRefs)
            {
                var templateField = template.SiteFields.FirstOrDefault(tf => tf.GetFieldId(parser) == fr.Id);
                if (templateField == null || templateField.GetFieldProvisioningStep(parser) == _step)
                {
                    fieldsRefsToProcess.Add(fr);
                }
            }

            foreach (var fieldRef in fieldsRefsToProcess)
            {
                Microsoft.SharePoint.Client.Field field = null;
                try
                {
                    field = web.AvailableFields.GetById(fieldRef.Id);
                }
                catch (ArgumentException)
                {
                    if (!string.IsNullOrEmpty(fieldRef.Name))
                    {
                        field = web.AvailableFields.GetByInternalNameOrTitle(fieldRef.Name);
                    }
                }
                // Add it to the target content type
                // Notice that this code will fail if the field does not exist
                web.AddFieldToContentType(createdCT, field,
                    fieldRef.Required, 
                    fieldRef.Hidden,
                    fieldRef.UpdateChildren
#if !SP2013 && !SP2016
                    ,
                    fieldRef.ShowInDisplayForm,
                    fieldRef.ReadOnly
#endif
                    );
            }

            // Add new CTs
            parser.AddToken(new ContentTypeIdToken(web, name, id));

#if !SP2013 && !SP2016
            // Set resources
            if (templateContentType.Name.ContainsResourceToken())
            {
                createdCT.NameResource.SetUserResourceValue(templateContentType.Name, parser);
            }
            if (templateContentType.Description.ContainsResourceToken())
            {
                createdCT.DescriptionResource.SetUserResourceValue(templateContentType.Description, parser);
            }
#endif
            //Reorder the elements so that the new created Content Type has the same order as defined in the
            //template. The order can be different if the new Content Type inherits from another Content Type.
            //In this case the new Content Type has all field of the original Content Type and missing fields
            //will be added at the end. To fix this issue we ordering the fields once more.

            var ctFields = templateContentType.FieldRefs.Select(fld => parser.ParseString(fld.Name)).ToArray();
            if (ctFields.Length > 0)
            {
                createdCT.FieldLinks.Reorder(ctFields);
            }
            // Set Hidden and Sealed property, ReadOnly will be set later
            if (createdCT.Hidden != templateContentType.Hidden)
            {
                createdCT.Hidden = templateContentType.Hidden;
            }
            if (createdCT.Sealed != templateContentType.Sealed)
            {
                createdCT.Sealed = templateContentType.Sealed;
            }

            if (templateContentType.DocumentSetTemplate == null)
            {
                // Only apply a document template when the contenttype is not a document set
                //Skipping updates of DocumentTemplate as we can't upload files to /_cts/ContentTypeName/FileName to noscript sites
                if (!isNoScriptSite)
                {
                    // Only apply a document template when the contenttype is not a document set
                    if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.DocumentTemplate)))
                    {
                        string documentTemplate = parser.ParseString(templateContentType.DocumentTemplate);
                        web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);
                        try
                        {
                            using (var fsstream = template.Connector.GetFileStream($"_cts/{name}/{documentTemplate}"))
                            {
                                if (fsstream != null)
                                {
                                    Microsoft.SharePoint.Client.Folder ctFolder = web.GetFolderByServerRelativeUrl($"{web.ServerRelativeUrl}/_cts/{name}");
                                    web.Context.Load(ctFolder, fl => fl.Files.Include(f => f.Name, f => f.ServerRelativeUrl));
                                    web.Context.ExecuteQueryRetry();

                                    FileCreationInformation newFile = new FileCreationInformation();
                                    newFile.ContentStream = fsstream;
                                    newFile.Url = $"{web.ServerRelativeUrl}/_cts/{name}/{documentTemplate}";

                                    Microsoft.SharePoint.Client.File uploadedFile = ctFolder.Files.Add(newFile);
                                    web.Context.Load(uploadedFile);
                                    web.Context.ExecuteQueryRetry();
                                }
                            }
                            createdCT.DocumentTemplate = documentTemplate;
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(ex, CoreResources.Provisioning_ObjectHandlers_ContentTypes_ErrorDocumentTemplate, name, documentTemplate);
                        }
                    }
                }
                else
                {
                    var parsedDocumentTemplate = parser.ParseString(templateContentType.DocumentTemplate);
                    if (!string.IsNullOrEmpty(parsedDocumentTemplate))
                    {
                        createdCT.DocumentTemplate = parsedDocumentTemplate;
                        // log message that's we are skipping uploads
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ContentTypes_SkipDocumentTemplate, name);
                    }
                }
            }

            // Skipping updates of forms as we can't upload forms to noscript sites
            if (!isNoScriptSite)
            {
                if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.NewFormUrl)))
                {
                    createdCT.NewFormUrl = parser.ParseString(templateContentType.NewFormUrl);
                }
                if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.EditFormUrl)))
                {
                    createdCT.EditFormUrl = parser.ParseString(templateContentType.EditFormUrl);
                }
                if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.DisplayFormUrl)))
                {
                    createdCT.DisplayFormUrl = parser.ParseString(templateContentType.DisplayFormUrl);
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.DisplayFormUrl)) ||
                    !string.IsNullOrEmpty(parser.ParseString(templateContentType.EditFormUrl)) ||
                    !string.IsNullOrEmpty(parser.ParseString(templateContentType.NewFormUrl)))
                {
                    // log message
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ContentTypes_SkipCustomFormUrls, name);
                }
            }

            createdCT.Update(true);
            web.Context.ExecuteQueryRetry();

            // If the CT is a DocumentSet
            if (templateContentType.DocumentSetTemplate != null)
            {
                // Retrieve a reference to the DocumentSet Content Type
                Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate documentSetTemplate =
                    Microsoft.SharePoint.Client.DocumentSet.DocumentSetTemplate.GetDocumentSetTemplate(web.Context, createdCT);

                // Load the collections to allow for deletion scenarions
                web.Context.Load(documentSetTemplate, d => d.AllowedContentTypes, d => d.DefaultDocuments, d => d.SharedFields, d => d.WelcomePageFields);
                web.Context.ExecuteQueryRetry();

                if (!String.IsNullOrEmpty(templateContentType.DocumentSetTemplate.WelcomePage))
                {
                    // TODO: Customize the WelcomePage of the DocumentSet
                }

                // Add additional content types to the set of allowed content types
                bool hasDefaultDocumentContentTypeInTemplate = false;
                foreach (String ctId in templateContentType.DocumentSetTemplate.AllowedContentTypes)
                {
                    Microsoft.SharePoint.Client.ContentType ct = existingCTs.FirstOrDefault(c => c.StringId == ctId);
                    if (ct != null)
                    {
                        if (ct.Id.StringValue.Equals("0x0101", StringComparison.InvariantCultureIgnoreCase))
                        {
                            hasDefaultDocumentContentTypeInTemplate = true;
                        }

                        documentSetTemplate.AllowedContentTypes.Add(ct.Id);
                    }
                }
                // If the default document content type (0x0101) is not in our definition then remove it
                if (!hasDefaultDocumentContentTypeInTemplate)
                {
                    Microsoft.SharePoint.Client.ContentType ct = existingCTs.FirstOrDefault(c => c.StringId == "0x0101");
                    if (ct != null)
                    {
                        documentSetTemplate.AllowedContentTypes.Remove(ct.Id);
                    }
                }

                if (!isNoScriptSite)
                {
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
                }
                else
                {
                    if (templateContentType.DocumentSetTemplate.DefaultDocuments.Any())
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ContentTypes_SkipDocumentSetDefaultDocuments, name);
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
            else if (templateContentType.Id.StartsWith(BuiltInContentTypeId.Workflow2013Task + "00"))
            {
                // If the Workflow Task (SP2013) contains more than one outcomeChoice, the Form UI will not show
                // the buttons associated each to choices, but fallback to classic Save and Cancel buttons.
                // +"00" is used to target only inherited content types and not alter OOB
                var outcomeFields = web.Context.LoadQuery(
                    createdCT.Fields.Where(f => f.TypeAsString == "OutcomeChoice"));
                web.Context.ExecuteQueryRetry();

                if (outcomeFields.Any())
                {
                    // 2 OutcomeChoice specified means the user has certainly push its own.
                    // Let's remove the default outcome field
                    var field = outcomeFields.FirstOrDefault(f => f.StaticName == "TaskOutcome");
                    if (field != null)
                    {
                        var fl = createdCT.FieldLinks.GetById(field.Id);
                        fl.DeleteObject();
                        createdCT.Update(true);
                        web.Context.ExecuteQueryRetry();
                    }
                }
            }

            web.Context.Load(createdCT);
            web.Context.ExecuteQueryRetry();

            return createdCT;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                template.ContentTypes.AddRange(GetEntities(web, scope, creationInfo, template));

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate, scope);
                }
            }
            return template;
        }

        private IEnumerable<ContentType> GetEntities(Web web, PnPMonitoredScope scope, ProvisioningTemplateCreationInformation creationInfo, ProvisioningTemplate template)
        {
            var cts = web.ContentTypes;
            web.Context.Load(cts, 
                ctCollection => ctCollection.IncludeWithDefaultProperties(
                    ct => ct.FieldLinks,
                    ct => ct.SchemaXmlWithResourceTokens
#if !SP2013 && !SP2016
                    ,
                    ct => ct.FieldLinks.IncludeWithDefaultProperties(
                        fl => fl.DisplayName,
                        fl => fl.ReadOnly, 
                        fl => fl.ShowInDisplayForm)
                    
#endif
                    )
                );

            web.Context.ExecuteQueryRetry();

            if (cts.Count > 0 && web.IsSubSite())
            {
                WriteMessage("We discovered content types in this subweb. While technically possible, we recommend moving these content types to the root site collection. Consider excluding them from this template.", ProvisioningMessageType.Warning);
            }
            web.EnsureProperties(w => w.Url, w => w.ServerRelativeUrl);//neded for DocumentTemplate extraction

            List<ContentType> ctsToReturn = new List<ContentType>();
            var currentCtIndex = 0;
            foreach (var ct in cts)
            {
                currentCtIndex++;
                WriteSubProgress("Content Type", ct.Name, currentCtIndex, cts.Count);

                if (!BuiltInContentTypeId.Contains(ct.StringId) &&
                    (creationInfo.ContentTypeGroupsToInclude.Count == 0 || creationInfo.ContentTypeGroupsToInclude.Contains(ct.Group)))
                {
                    // Exclude the content type if it's from syndication, and if the flag is not set
                    if (!creationInfo.IncludeContentTypesFromSyndication && IsContentTypeFromSyndication(ct))
                    {
                        scope.LogInfo($"Content type {ct.Name} excluded from export because it's a syndicated content type.");

                        continue;
                    }

                    string ctDocumentTemplate = null;
                    if (!string.IsNullOrEmpty(ct.DocumentTemplate))
                    {
                        if (!ct.DocumentTemplate.StartsWith("_cts/"))
                        {
                            ctDocumentTemplate = ct.DocumentTemplate;
                        }
                        //extract DocumentTemplate if it points to ContentType Ressource Folder
                        if (creationInfo.ExtractConfiguration != null && creationInfo.ExtractConfiguration.PersistAssetFiles && !string.IsNullOrWhiteSpace(ct.DocumentTemplateUrl) && ct.DocumentTemplateUrl.Contains("_cts/"))
                        {
                            try
                            {
                                var spFile = web.GetFileByServerRelativeUrl(ct.DocumentTemplateUrl);
                                spFile.EnsureProperties(f => f.Level, f => f.ServerRelativeUrl, f => f.Name);

                                // If we got here it's a file, let's grab the file's path and name
                                var baseUri = new Uri(web.Url);
                                var fullUri = new Uri(baseUri, spFile.ServerRelativeUrl);
                                var folderPath = System.Web.HttpUtility.UrlDecode(fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
                                var fileName = System.Web.HttpUtility.UrlDecode(fullUri.Segments[fullUri.Segments.Count() - 1]);

                                var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

                                web.Context.Load(spFile);
                                web.Context.ExecuteQueryRetry();
                                var spFileStream = spFile.OpenBinaryStream();
                                web.Context.ExecuteQueryRetry();

                                template.Connector.SaveFileStream(spFile.Name, templateFolderPath, spFileStream.Value);
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(ex, CoreResources.Provisioning_ObjectHandlers_ContentTypes_ErrorSaveDocumentTemplateToConnector, ct.Name, ct.DocumentTemplateUrl);
                            }
                        }
                    }

                    var newCT = new ContentType(
                        ct.StringId,
                        ct.Name,
                        ct.Description,
                        ct.Group,
                        ct.Sealed,
                        ct.Hidden,
                        ct.ReadOnly,
                        ctDocumentTemplate,
                        false,
                            (from fieldLink in ct.FieldLinks.AsEnumerable<FieldLink>()
                             select new FieldRef(fieldLink.Name)
                             {
                                 Id = fieldLink.Id,
                                 Hidden = fieldLink.Hidden,
                                 Required = fieldLink.Required,
#if !SP2013 && !SP2016
                                 DisplayName = fieldLink.DisplayName,
                                 ShowInDisplayForm = fieldLink.ShowInDisplayForm,
                                 ReadOnly = fieldLink.ReadOnly,
#endif
                             })
                        )
                    {
                        DisplayFormUrl = ct.DisplayFormUrl,
                        EditFormUrl = ct.EditFormUrl,
                        NewFormUrl = ct.NewFormUrl,
                    };

                    if (creationInfo.PersistMultiLanguageResources)
                    {
#if !SP2013
                        // only persist language values for content types we actually will keep...no point in spending time on this is we clean the field afterwards
                        var persistLanguages = true;
                        if (creationInfo.BaseTemplate != null)
                        {
                            int index = creationInfo.BaseTemplate.ContentTypes.FindIndex(c => c.Id.Equals(ct.StringId));

                            if (index > -1)
                            {
                                persistLanguages = false;
                            }
                        }

                        if (persistLanguages)
                        {
                            var escapedCTName = ct.Name.Replace(" ", "_");
                            if (UserResourceExtensions.PersistResourceValue(ct.NameResource, $"ContentType_{escapedCTName}_Title", template, creationInfo))
                            {
                                newCT.Name = $"{{res:ContentType_{escapedCTName}_Title}}";
                            }
                            if (UserResourceExtensions.PersistResourceValue(ct.DescriptionResource, $"ContentType_{escapedCTName}_Description", template, creationInfo))
                            {
                                newCT.Description = $"{{res:ContentType_{escapedCTName}_Description}}";
                            }
                        }
#endif
                    }

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
                             select allowedCT.StringValue).ToList(),
                            (from defaultDocument in documentSetTemplate.DefaultDocuments.AsEnumerable()
                             select new DefaultDocument
                             {
                                 ContentTypeId = defaultDocument.ContentTypeId.StringValue,
                                 Name = defaultDocument.Name,
#if SP2013 || SP2016
								 FileSourcePath = string.Empty
#else
                                 FileSourcePath = creationInfo.PersistBrandingFiles ? $"_cts/{ct.Name}/{defaultDocument.DocumentPath.DecodedUrl}" : string.Empty
#endif
                             }).ToList(),
                            (from sharedField in documentSetTemplate.SharedFields.AsEnumerable()
                             select sharedField.Id).ToList(),
                            (from welcomePageField in documentSetTemplate.WelcomePageFields.AsEnumerable()
                             select welcomePageField.Id).ToList()
                        );

                        //extract the DefaultDocument files
                        foreach (var defaultDoc in newCT.DocumentSetTemplate.DefaultDocuments.Where(dd => !string.IsNullOrWhiteSpace(dd.FileSourcePath)))
                        {
                            try
                            {
                                string serverRelativeUrl = $"{web.ServerRelativeUrl}/{defaultDoc.FileSourcePath}";
                                var spFile = web.GetFileByServerRelativeUrl(serverRelativeUrl);

                                spFile.EnsureProperties(f => f.Level, f => f.ServerRelativeUrl, f => f.Name);

                                // If we got here it's a file, let's grab the file's path and name
                                var baseUri = new Uri(web.Url);
                                var fullUri = new Uri(baseUri, spFile.ServerRelativeUrl);
                                var folderPath = System.Web.HttpUtility.UrlDecode(fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
                                var fileName = System.Web.HttpUtility.UrlDecode(fullUri.Segments[fullUri.Segments.Count() - 1]);

                                var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

                                web.Context.Load(spFile);
                                web.Context.ExecuteQueryRetry();
                                var spFileStream = spFile.OpenBinaryStream();
                                web.Context.ExecuteQueryRetry();

                                template.Connector.SaveFileStream(spFile.Name, templateFolderPath, spFileStream.Value);
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(ex, CoreResources.Provisioning_ObjectHandlers_ContentTypes_ErrorExtractDocumentSetTemplate, defaultDoc.FileSourcePath);
                            }
                        }
                    }

                    ctsToReturn.Add(newCT);
                }
            }
            WriteMessage("Done processing Content Types", ProvisioningMessageType.Completed);
            return ctsToReturn;
        }

        private static bool IsContentTypeFromSyndication(Microsoft.SharePoint.Client.ContentType ct)
        {
            if (ct == null) throw new ArgumentNullException(nameof(ct));

            var schema = XElement.Parse(ct.SchemaXmlWithResourceTokens);
            var xmlNsMgr = new XmlNamespaceManager(new NameTable());
            xmlNsMgr.AddNamespace("cts", "Microsoft.SharePoint.Taxonomy.ContentTypeSync");
            var contentTypeSyncB64 = schema.XPathSelectElement("/XmlDocuments/XmlDocument[@NamespaceURI='Microsoft.SharePoint.Taxonomy.ContentTypeSync']", xmlNsMgr)?.Value;

            if (contentTypeSyncB64 == null) return false;

            var contentTypeSyncString = Encoding.UTF8.GetString(Convert.FromBase64String(contentTypeSyncB64));

            var contentTypeXml = XElement.Parse(contentTypeSyncString);

            // If id is different, that means the ContentTypeSync document is inherited from its parent.
            // That means this content type is not syndicated
            return (string)contentTypeXml.Attribute("ContentTypeId") == ct.Id.StringValue;
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

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
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
            using (var mem = new MemoryStream())
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