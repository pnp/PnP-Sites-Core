using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Interface to test if a template can be provisioned onto a target tenant
    /// </summary>
    internal interface ICanProvisionRuleTenant: ICanProvisionRuleBase
    {
        /// <summary>
        /// This method allows to check if a template can be provisioned in the currently selected target
        /// </summary>
        /// <param name="tenant">The target Tenant</param>
        /// <param name="hierarchy">The Template to hierarchy</param>
        /// <param name="sequenceId">The sequence to test within the hierarchy</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        CanProvisionResult CanProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation);
    }
}
