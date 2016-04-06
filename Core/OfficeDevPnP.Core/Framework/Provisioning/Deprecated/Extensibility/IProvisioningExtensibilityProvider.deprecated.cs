using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    /// Defines a interface that accepts requests from the provisioning processing component
    /// </summary>
    [Obsolete("Use IProvisioningExtensibilityHandler")]
    public interface IProvisioningExtensibilityProvider
    {
        /// <summary>
        /// Defines a interface that accepts requests from the provisioning processing component
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="template"></param>
        /// <param name="configurationData"></param>
        void ProcessRequest(ClientContext ctx, ProvisioningTemplate template, string configurationData);

    }
}
