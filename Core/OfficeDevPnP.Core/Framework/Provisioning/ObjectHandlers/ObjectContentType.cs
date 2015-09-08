using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using ContentType = OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;

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

                web.Context.Load(web.ContentTypes, ct => ct.IncludeWithDefaultProperties(c => c.StringId, c => c.FieldLinks));
                web.Context.ExecuteQueryRetry();
                var existingCTs = web.ContentTypes.ToList();

                foreach (var ct in template.ContentTypes.OrderBy(ct => ct.Id)) // ordering to handle references to parent content types that can be in the same template
                {
                    var existingCT = existingCTs.FirstOrDefault(c => c.StringId.Equals(ct.Id, StringComparison.OrdinalIgnoreCase));
                    if (existingCT == null)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_ContentTypes_Creating_new_Content_Type___0_____1_, ct.Id, ct.Name);
                        var newCT = CreateContentType(web, ct, parser);
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
                            var newCT = CreateContentType(web, ct, parser);
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
            }
            if (templateContentType.Group != null && existingContentType.Group != parser.ParseString(templateContentType.Group))
            {
                scope.LogPropertyUpdate("Group");
                existingContentType.Group = parser.ParseString(templateContentType.Group);
                isDirty = true;
            }
            if (isDirty)
            {
                existingContentType.Update(true);
                web.Context.ExecuteQueryRetry();
            }
            // Delta handling
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
            if (isDirty)
            {
                existingContentType.Update(true);
                web.Context.ExecuteQueryRetry();
            }
        }

        private static Microsoft.SharePoint.Client.ContentType CreateContentType(Web web, ContentType templateContentType, TokenParser parser)
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

            createdCT.ReadOnly = templateContentType.ReadOnly;
            createdCT.Hidden = templateContentType.Hidden;
            createdCT.Sealed = templateContentType.Sealed;
            if (!string.IsNullOrEmpty(parser.ParseString(templateContentType.DocumentTemplate)))
            {
                createdCT.DocumentTemplate = parser.ParseString(templateContentType.DocumentTemplate);
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
                    ctsToReturn.Add(new ContentType
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
                        ));
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
    }
}
