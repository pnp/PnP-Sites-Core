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
    [CanProvisionRule(Scope = CanProvisionScope.Site, Sequence = 100)]
    internal class CanProvisionAppsRuleSite : CanProvisionRuleSiteBase
    {
        public override CanProvisionResult CanProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {

            // Prepare the default output
            var result = new CanProvisionResult();
#if !ONPREMISES
            // Verify if we need the App Catalog (i.e. the template contains apps or packages)
            if ((template.ApplicationLifecycleManagement?.Apps != null && template.ApplicationLifecycleManagement?.Apps?.Count > 0) ||
                template.ApplicationLifecycleManagement?.AppCatalog != null)
            {
                using (var scope = new PnPMonitoredScope(this.Name))
                {
                    // Try to access the AppCatalog
                    var appCatalogUri = web.GetAppCatalog();
                    if (appCatalogUri == null)
                    {
                        // And if we fail, raise a CanProvisionIssue
                        result.CanProvision = false;
                        result.Issues.Add(new CanProvisionIssue()
                        {
                            Source = this.Name,
                            Tag = CanProvisionIssueTags.MISSING_APP_CATALOG,
                            Message = CanProvisionIssuesMessages.Missing_App_Catalog,
                            InnerException = null, // Here we don't have any specific exception
                        });
                    }
                }
            }
#else
            result.CanProvision = false;
#endif
            return result;
        }
    }
}
