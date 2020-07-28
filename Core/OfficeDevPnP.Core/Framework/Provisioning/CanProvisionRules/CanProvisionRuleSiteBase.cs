using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Base class to test if the Provisioning Template can be provisioned on the target SharePoint Site
    /// </summary>
    internal abstract class CanProvisionRuleSiteBase: ICanProvisionRuleSite
    {
        /// <summary>
        /// The TenantAdminSiteUrl contains the tenant admin site url. 
        /// <remarks>
        /// This value is only relevant for onpremesis environments (SP2013, SP2016, SP2019)
        /// </remarks>
        /// </summary>
        public string TenantAdminSiteUrl { get; set; }

        public string Name { get => this.GetType().FullName; }

        /// <summary>
        /// This method allows to check if a template can be provisioned in the currently selected target
        /// </summary>
        /// <param name="web">The target Web</param>
        /// <param name="template">The Template to provision</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        public virtual CanProvisionResult CanProvision(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // By default everything can be provisioned
            return (new CanProvisionResult { CanProvision = true, Issues = null });
        }
    }
}
