using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#if !NETSTANDARD2_0
namespace OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut
{
    public class ExtensibilityMockTokenProvider : IProvisioningExtensibilityTokenProvider
    {
        public static ClientContext ReceivedCtx = null;
        public static ProvisioningTemplate ReceivedProvisioningTemplate = null;
        public static string ReceivedConfigurationData = null;

        public IEnumerable<TokenDefinition> GetTokens(ClientContext ctx, ProvisioningTemplate template, string configurationData)
        {
            ExtensibilityMockTokenProvider.ReceivedCtx = ctx;
            ExtensibilityMockTokenProvider.ReceivedProvisioningTemplate = template;
            ExtensibilityMockTokenProvider.ReceivedConfigurationData = configurationData;

            return new List<TokenDefinition> { new MockToken(ctx.Web) };
        }
    }

    public class MockToken : TokenDefinition
    {
        public const string MockTokenKey = "{mocktoken}";
        public const string MockTokenValue = "ValueFromMockToken";

        public MockToken(Web web) : base(web, MockTokenKey)
        {
        }
        
        public override string GetReplaceValue()
        {
            return MockTokenValue;
        }
    }
}
#endif