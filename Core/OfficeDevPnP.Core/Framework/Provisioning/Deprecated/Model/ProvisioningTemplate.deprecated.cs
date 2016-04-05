using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object for the Provisioning Template
    /// </summary>
    public partial class ProvisioningTemplate
    {
        private ProviderCollection _providers;

        /// <summary>
        /// Gets a collection of Providers that are used during the extensibility pipeline
        /// </summary>
        /// 
        [Obsolete("Use ExtensibilityHandlers")]
        public ProviderCollection Providers
        {
            get { return this._providers; }
            private set { this._providers = value; }
        }
    }
}
