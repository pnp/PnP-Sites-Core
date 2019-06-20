using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules.Rules
{
    [CanProvisionRule(Scope = CanProvisionScope.Tenant, Sequence = 200)]
    internal class CanProvisionTermStoreRuleTenant : CanProvisionRuleTenantBase
    {
        public override CanProvisionResult CanProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // Rely on the corresponding Site level CanProvision rule
            return (this.EvaluateSiteRule<CanProvisionTermStoreRuleSite>(tenant, hierarchy, sequenceId, applyingInformation));
        }
    }
}
