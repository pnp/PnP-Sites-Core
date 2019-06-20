using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    /// Defines a interface that accepts requests from the provisioning processing component
    /// </summary>
    [Obsolete("Use IProvisioningExtensibilityHandler. This method will be removed in the June 2016 release.")]
    public interface IProvisioningExtensibilityProvider
    {
        /// <summary>
        /// Defines a interface that accepts requests from the provisioning processing component
        /// </summary>
        /// <param name="ctx">The ClientContext to process</param>
        /// <param name="template">The Provisioning template</param>
        /// <param name="configurationData">Configuration Data string</param>
        void ProcessRequest(ClientContext ctx, ProvisioningTemplate template, string configurationData);

    }
}
