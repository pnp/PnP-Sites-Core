using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules.Rules
{
    [CanProvisionRule(Scope = CanProvisionScope.Tenant, Sequence = 100)]
    internal class CanProvisionAppsRuleTenant : CanProvisionRuleTenantBase
    {
        public override CanProvisionResult CanProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // Rely on the corresponding Site level CanProvision rule
            return (this.EvaluateSiteRule<CanProvisionAppsRuleSite>(tenant, hierarchy, sequenceId, applyingInformation));
        }
    }
}
