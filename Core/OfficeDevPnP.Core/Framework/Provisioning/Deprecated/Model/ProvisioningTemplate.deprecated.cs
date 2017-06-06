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
        [Obsolete("Use ExtensibilityHandlers. This property will be removed in the June 2016 release.")]
        public ProviderCollection Providers
        {
            get { return this._providers; }
            private set { this._providers = value; }
        }

        /// <summary>
        /// The Search Settings for the Provisioning Template
        /// </summary>
        [Obsolete("Use SiteSearchSettings or WebSearchSettings. This property will be removed in the September 2016 release.")]
        public String SearchSettings
        {
            get { return this.SiteSearchSettings; }
            set { this.SiteSearchSettings = value; }
        }
    }
}
