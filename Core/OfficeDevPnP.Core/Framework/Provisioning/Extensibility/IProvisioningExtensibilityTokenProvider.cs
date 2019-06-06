using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    /// Defines an interface which allows to plugin custom TokenDefinitions to the template provisioning pipeline
    /// </summary>
    public interface IProvisioningExtensibilityTokenProvider
    {
        /// <summary>
        /// Provides Token Definitions to the template provisioning pipeline
        /// </summary>
        /// <param name="ctx">The ClientContext</param>
        /// <param name="template">The Provisioning template</param>
        /// <param name="configurationData">Configuration Data string</param>
        IEnumerable<TokenDefinition> GetTokens(ClientContext ctx, ProvisioningTemplate template, string configurationData);
    }
}
