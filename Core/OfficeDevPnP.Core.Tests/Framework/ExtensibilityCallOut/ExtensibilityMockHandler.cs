using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

#if !NETSTANDARD2_0
namespace OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut
{
    public class ExtensibilityMockHandler : IProvisioningExtensibilityHandler
    {
        public ProvisioningTemplate Extract(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation, PnPMonitoredScope scope, string configurationData)
        {
            template.Lists.Add(new ListInstance() { Title = "Test List" });

            return template;
        }

        public IEnumerable<TokenDefinition> GetTokens(ClientContext ctx, ProvisioningTemplate template, string configurationData)
        {
            throw new NotImplementedException();
        }

        public void Provision(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation, TokenParser tokenParser, PnPMonitoredScope scope, string configurationData)
        {
            bool _urlCheck = ctx.Url.Equals(ExtensibilityTestConstants.MOCK_URL, StringComparison.OrdinalIgnoreCase);
            if (!_urlCheck) throw new Exception("CTXURLNOTTHESAME");

            bool _templateCheck = template.Id.Equals(ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID, StringComparison.OrdinalIgnoreCase);
            if (!_templateCheck) throw new Exception("TEMPLATEIDNOTTHESAME");

            bool _configDataCheck = configurationData.Equals(ExtensibilityTestConstants.PROVIDER_MOCK_DATA, StringComparison.OrdinalIgnoreCase);
            if (!_configDataCheck) throw new Exception("CONFIGDATANOTTHESAME");
        }
    }
}
#endif