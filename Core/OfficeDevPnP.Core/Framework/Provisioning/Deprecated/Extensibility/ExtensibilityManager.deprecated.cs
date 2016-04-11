using System;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    ///  Provisioning Framework Component that is used for invoking custom providers during the provisioning process.
    /// </summary>
    public partial class ExtensibilityManager
    {

        /// <summary>
        /// Method to Invoke Custom Provisioning Providers. 
        /// Ensure the ClientContext is not disposed in the custom provider.
        /// </summary>
        /// <param name="ctx">Authenticated ClientContext that is passed to the custom provider.</param>
        /// <param name="handler">A custom Extensibility Provisioning Provider</param>
        /// <param name="template">ProvisioningTemplate that is passed to the custom provider</param>
        /// <exception cref="ExtensiblityPipelineException"></exception>
        /// <exception cref="ArgumentException">Provider.Assembly or Provider.Type is NullOrWhiteSpace></exception>
        /// <exception cref="ArgumentNullException">ClientContext is Null></exception>
        [Obsolete("Use ExecuteExtensibilityProvisionCallOut. This method will be removed in the June 2016 release.")]
        public void ExecuteExtensibilityCallOut(ClientContext ctx, ExtensibilityHandler handler, ProvisioningTemplate template)
        {
            ExecuteExtensibilityProvisionCallOut(ctx, handler, template, null, null, null);
        }

    }
}
