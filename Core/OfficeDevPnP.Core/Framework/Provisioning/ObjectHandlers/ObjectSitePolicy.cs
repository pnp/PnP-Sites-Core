using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSitePolicy : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Site Policy"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_SitePolicy))
            {
                if (template.SitePolicy != null)
                {
                    if (web.GetSitePolicyByName(template.SitePolicy) != null) // Site Policy Available?
                    {
                        web.ApplySitePolicy(template.SitePolicy);
                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_SitePolicy))
            {
                var sitePolicyEntity = web.GetAppliedSitePolicy();

                if (sitePolicyEntity != null)
                {
                    template.SitePolicy = sitePolicyEntity.Name;
                }
            }
            return template;
        }


        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.SitePolicy != null;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                var sitePolicyEntity = web.GetAppliedSitePolicy();

                _willExtract = sitePolicyEntity != null;
            }
            return _willExtract.Value;
        }
    }
}

