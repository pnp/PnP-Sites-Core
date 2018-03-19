using System;
using System.Linq;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Field = Microsoft.SharePoint.Client.Field;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectLookupFields : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Lookup Fields"; }
        }

        public ObjectLookupFields()
        {
            this.ReportProgress = false;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                try
                {
                    parser = ProcessLookupFields(web, template, parser, scope);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_LookupFields_Processing_lookup_fields_failed___0_____1_, ex.Message, ex.StackTrace);
                    throw;
                }

            }

            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            //using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_LookupFields))
            //{ }
            return template;
        }

        private TokenParser ProcessLookupFields(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope)
        {
            var rootWeb = (web.Context as ClientContext).Site.RootWeb;
            rootWeb.Context.Load(rootWeb.Lists, lists => lists.Include(l => l.Id, l => l.RootFolder.ServerRelativeUrl, l => l.Fields).Where(l => l.Hidden == false));
            rootWeb.Context.ExecuteQueryRetry();

            foreach (var siteField in template.SiteFields)
            {
                var fieldElement = XElement.Parse(siteField.SchemaXml);

                if (fieldElement.Attribute("List") != null)
                {
                    var fieldId = Guid.Parse(fieldElement.Attribute("ID").Value);
                    var listIdentifier = parser.ParseString(fieldElement.Attribute("List").Value);
                    var relationshipDeleteBehavior = fieldElement.Attribute("RelationshipDeleteBehavior") != null ? fieldElement.Attribute("RelationshipDeleteBehavior").Value : string.Empty;
                    var webId = string.Empty;

                    var field = rootWeb.Fields.GetById(fieldId);
                    rootWeb.Context.Load(field, f => f.SchemaXmlWithResourceTokens);
                    rootWeb.Context.ExecuteQueryRetry();

                    List sourceList = FindSourceList(listIdentifier, web, rootWeb);

                    if (sourceList != null)
                    {
                        rootWeb.Context.Load(sourceList.ParentWeb);
                        rootWeb.Context.ExecuteQueryRetry();

                        webId = sourceList.ParentWeb.Id.ToString();

                        ProcessField(field, sourceList.Id, webId, relationshipDeleteBehavior);
                    }
                }
            }

            web.Context.Load(web.Lists, lists => lists.Include(l => l.Id, l => l.RootFolder.ServerRelativeUrl, l => l.Fields).Where(l => l.Hidden == false));
            web.Context.ExecuteQueryRetry();

            foreach (var listInstance in template.Lists)
            {
                foreach (var listField in listInstance.Fields)
                {
                    var fieldElement = XElement.Parse(listField.SchemaXml);
                    if (fieldElement.Attribute("List") == null) continue;

                    var fieldId = Guid.Parse(fieldElement.Attribute("ID").Value);
                    var listIdentifier = parser.ParseString(fieldElement.Attribute("List").Value);
                    var relationshipDeleteBehavior = fieldElement.Attribute("RelationshipDeleteBehavior") != null ? fieldElement.Attribute("RelationshipDeleteBehavior").Value : string.Empty;
                    var webId = string.Empty;

                    var listUrl = UrlUtility.Combine(web.ServerRelativeUrl, parser.ParseString(listInstance.Url));

                    var createdList = web.Lists.FirstOrDefault(l => l.RootFolder.ServerRelativeUrl.Equals(listUrl, StringComparison.OrdinalIgnoreCase));
                    if (createdList != null)
                    {
                        try
                        {
                            var field = createdList.Fields.GetById(fieldId);
                            web.Context.Load(field, f => f.SchemaXmlWithResourceTokens);
                            web.Context.ExecuteQueryRetry();

                            List sourceList = FindSourceList(listIdentifier, web, rootWeb);

                            if (sourceList != null)
                            {
                                web.Context.Load(sourceList.ParentWeb);
                                web.Context.ExecuteQueryRetry();

                                webId = sourceList.ParentWeb.Id.ToString();
                                ProcessField(field, sourceList.Id, webId, relationshipDeleteBehavior);
                            }
                        }
                        catch (ArgumentException ex)
                        {
                            // We skip and log any issues related to not existing lookup fields
                            scope.LogError($"Exception searching for field! {ex.Message}");
                        }
                    }
                }
            }

            return parser;
        }

        private static List FindSourceList(string listIdentifier, Web web, Web rootWeb)
        {
            Guid listGuid = Guid.Empty;

            if (!Guid.TryParse(listIdentifier, out listGuid))
            {
                var sourceListUrl = UrlUtility.Combine(web.ServerRelativeUrl, (listIdentifier == Constants.FIELD_XML_USER_LISTIDENTIFIER ? Constants.FIELD_XML_USER_LISTRELATIVEURL : listIdentifier));
                return web.Lists.FirstOrDefault(l => l.RootFolder.ServerRelativeUrl.Equals(sourceListUrl, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                List retVal = rootWeb.Lists.FirstOrDefault(l => l.Id.Equals(listGuid));

                if(retVal == null)
                {
                    retVal = web.Lists.FirstOrDefault(l => l.Id.Equals(listGuid));
                }

                if(retVal == null)
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.Provisioning_ObjectHandlers_LookupFields_LookupTargetListLookupFailed__0, listIdentifier);
                }
                return retVal;
            }
        }

        private static void ProcessField(Field field, Guid listGuid, string webId, string relationshipDeleteBehavior)
        {
            var isDirty = false;

            var existingFieldElement = XElement.Parse(field.SchemaXmlWithResourceTokens);

            isDirty = UpdateFieldAttribute(existingFieldElement, "List", listGuid.ToString("B"), false);

            isDirty = UpdateFieldAttribute(existingFieldElement, "WebId", webId, isDirty);

            var webIdGuid = Guid.Parse(webId);

            isDirty = UpdateFieldAttribute(existingFieldElement, "SourceID", webIdGuid.ToString("B"), isDirty);

            if (!string.IsNullOrEmpty(relationshipDeleteBehavior))
                isDirty = UpdateFieldAttribute(existingFieldElement, "RelationshipDeleteBehavior", relationshipDeleteBehavior, isDirty);

            if (isDirty)
            {
                field.SchemaXml = existingFieldElement.ToString();

                field.UpdateAndPushChanges(true);
                field.Context.ExecuteQueryRetry();
            }
        }

        private static bool UpdateFieldAttribute(XElement existingFieldElement, string attributeName, string attributeValue, bool isDirty)
        {
            if (existingFieldElement.Attribute(attributeName) == null)
            {
                existingFieldElement.Add(new XAttribute(attributeName, attributeValue));
                isDirty = true;
            }
            else if (!existingFieldElement.Attribute(attributeName).Value.Equals(attributeValue))
            {
                existingFieldElement.Attribute(attributeName).SetValue(attributeValue);
                isDirty = true;
            }
            return isDirty;
        }


        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = true;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }
}
