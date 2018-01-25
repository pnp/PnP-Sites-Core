using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSearchSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Search Settings"; }
        }
        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var site = (web.Context as ClientContext).Site;
                try
                {
                    var siteSearchSettings = site.GetSearchConfiguration();

                    if (!String.IsNullOrEmpty(siteSearchSettings))
                    {
                        template.SiteSearchSettings = siteSearchSettings;
                    }

                    var webSearchSettings = web.GetSearchConfiguration();

                    if (!String.IsNullOrEmpty(webSearchSettings))
                    {
                        template.WebSearchSettings = webSearchSettings;
                    }
                }
                catch (ServerException)
                {
                    // The search service is not necessarily configured
                    // Swallow the exception
                }
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var site = (web.Context as ClientContext).Site;
                if (!String.IsNullOrEmpty(template.SiteSearchSettings))
                {
                    site.SetSearchConfiguration(template.SiteSearchSettings);
                }

                if (!String.IsNullOrEmpty(template.WebSearchSettings))
                {
                    web.SetSearchConfiguration(template.WebSearchSettings);
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return creationInfo.IncludeSearchConfiguration;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#pragma warning disable 618
            return !String.IsNullOrEmpty(template.SearchSettings);
#pragma warning restore 618
        }
    }
}
