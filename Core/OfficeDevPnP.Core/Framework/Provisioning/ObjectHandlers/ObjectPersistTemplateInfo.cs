using System;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPersistTemplateInfo : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Persist Template Info"; }
        }

        public ObjectPersistTemplateInfo()
        {
            this.ReportProgress = false;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.SetPropertyBagValue("_PnP_ProvisioningTemplateId", template.Id != null ? template.Id : "");
                web.AddIndexedPropertyBagKey("_PnP_ProvisioningTemplateId");

                ProvisioningTemplateInfo info = new ProvisioningTemplateInfo();
                info.TemplateId = template.Id != null ? template.Id : "";
                info.TemplateVersion = template.Version;
                info.TemplateSitePolicy = template.SitePolicy;
                info.Result = true;
                info.ProvisioningTime = DateTime.Now;

                string jsonInfo = JsonConvert.SerializeObject(info);

                web.SetPropertyBagValue("_PnP_ProvisioningTemplateInfo", jsonInfo);
            }
            return parser;
        }

        public override Model.ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            //using (var scope = new PnPMonitoredScope(this.Name))
            //{ }
            return template;
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
