using System;
using System.Linq;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Field = Microsoft.SharePoint.Client.Field;
using OfficeDevPnP.Core.Diagnostics;

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
                    var listIdentifier = fieldElement.Attribute("List").Value;
                    var webId = string.Empty;

                    var field = rootWeb.Fields.GetById(fieldId);
                    rootWeb.Context.Load(field, f => f.SchemaXml);
                    rootWeb.Context.ExecuteQueryRetry();

                    Guid listGuid;
                    if (!Guid.TryParse(listIdentifier, out listGuid))
                    {
                        var sourceListUrl = UrlUtility.Combine(web.ServerRelativeUrl, parser.ParseString(listIdentifier));
                        var sourceList = rootWeb.Lists.FirstOrDefault(l => l.RootFolder.ServerRelativeUrl.Equals(sourceListUrl, StringComparison.OrdinalIgnoreCase));
                        if (sourceList != null)
                        {
                            listGuid = sourceList.Id;

                            rootWeb.Context.Load(sourceList.ParentWeb);
                            rootWeb.Context.ExecuteQueryRetry();

                            webId = sourceList.ParentWeb.Id.ToString();
                        }
                    }
                    if (listGuid != Guid.Empty)
                    {
                        ProcessField(field, listGuid, webId);
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
                    var listIdentifier = fieldElement.Attribute("List").Value;
                    var webId = string.Empty;

                    var listUrl = UrlUtility.Combine(web.ServerRelativeUrl, parser.ParseString(listInstance.Url));

                    var createdList = web.Lists.FirstOrDefault(l => l.RootFolder.ServerRelativeUrl.Equals(listUrl, StringComparison.OrdinalIgnoreCase));
                    if (createdList != null)
                    {
                        var field = createdList.Fields.GetById(fieldId);
                        web.Context.Load(field, f => f.SchemaXml);
                        web.Context.ExecuteQueryRetry();

                        Guid listGuid;
                        if (!Guid.TryParse(parser.ParseString(listIdentifier), out listGuid))
                        {
                            var sourceListUrl = UrlUtility.Combine(web.ServerRelativeUrl, parser.ParseString(listIdentifier));
                            var sourceList = web.Lists.FirstOrDefault(l => l.RootFolder.ServerRelativeUrl.Equals(sourceListUrl, StringComparison.OrdinalIgnoreCase));
                            if (sourceList != null)
                            {
                                listGuid = sourceList.Id;

                                web.Context.Load(sourceList.ParentWeb);
                                web.Context.ExecuteQueryRetry();

                                webId = sourceList.ParentWeb.Id.ToString();
                            }
                        }
                        if (listGuid != Guid.Empty)
                        {
                            ProcessField(field, listGuid, webId);
                        }
                    }
                }
            }

            return parser;
        }

        private static void ProcessField(Field field, Guid listGuid, string webId)
        {
            var isDirty = false;

            var existingFieldElement = XElement.Parse(field.SchemaXml);

            isDirty = UpdateFieldAttribute(existingFieldElement, "List", listGuid.ToString(), false);

            isDirty = UpdateFieldAttribute(existingFieldElement, "WebId", webId, isDirty);

            isDirty = UpdateFieldAttribute(existingFieldElement, "SourceID", webId, isDirty);

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


        public override bool WillProvision(Web web, ProvisioningTemplate template)
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
