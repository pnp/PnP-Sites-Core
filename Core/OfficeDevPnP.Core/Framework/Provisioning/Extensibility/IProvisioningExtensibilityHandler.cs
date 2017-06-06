using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    /// Defines an interface which allows to plugin custom Provisioning Extensibility Handlers to the template extraction/provisioning pipeline
    /// </summary>
    public interface IProvisioningExtensibilityHandler : IProvisioningExtensibilityTokenProvider
    {
        /// <summary>
        /// Execute custom actions during provisioning of a template
        /// </summary>
        /// <param name="ctx">The target ClientContext</param>
        /// <param name="template">The current Provisioning Template</param>
        /// <param name="applyingInformation">The Provisioning Template application information object</param>
        /// <param name="tokenParser">Token parser instance</param>
        /// <param name="scope">The PnPMonitoredScope of the current step in the pipeline</param>
        /// <param name="configurationData">The configuration data, if any, for the handler</param>
        void Provision(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation, TokenParser tokenParser, PnPMonitoredScope scope, string configurationData);

        /// <summary>
        /// Execute custom actions during extraction of a template
        /// </summary>
        /// <param name="ctx">The target ClientContext</param>
        /// <param name="template">The current Provisioning Template</param>
        /// <param name="creationInformation">The Provisioning Template creation information object</param>
        /// <param name="scope">The PnPMonitoredScope of the current step in the pipeline</param>
        /// <param name="configurationData">The configuration data, if any, for the handler</param>
        /// <returns>The Provisioning Template eventually enriched by the handler during extraction</returns>
        ProvisioningTemplate Extract(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation, PnPMonitoredScope scope, string configurationData);
    }
}
