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
                    var searchSettings = site.GetSearchConfiguration();

                    if (!String.IsNullOrEmpty(searchSettings))
                    {
                        template.SearchSettings = searchSettings;
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
                if (!String.IsNullOrEmpty(template.SearchSettings))
                {
                    site.SetSearchConfiguration(template.SearchSettings);
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return creationInfo.IncludeSearchConfiguration;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return !String.IsNullOrEmpty(template.SearchSettings);
        }
    }
}
