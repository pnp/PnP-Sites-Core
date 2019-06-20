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
    /// Base abstract class to test if the Provisioning Template can be provisioned on the target Office 365 Tenant
    /// </summary>
    internal abstract class CanProvisionRuleOffice365Base: ICanProvisionRuleOffice365
    {
        public string Name { get => this.GetType().FullName; }

        public int Sequence { get => 0; }

        /// <summary>
        /// This method allows to check if a template can be provisioned
        /// </summary>
        /// <param name="hierarchy">The Template to hierarchy</param>
        /// <param name="sequenceId">The sequence to test within the hierarchy</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        public virtual CanProvisionResult CanProvision(Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // By default everything can be provisioned
            return (new CanProvisionResult { CanProvision = true, Issues = null });
        }
    }
}
