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
    /// Base class to test if the Provisioning Template can be provisioned on the target SharePoint Tenant
    /// </summary>
    internal abstract class CanProvisionRuleTenantBase: ICanProvisionRuleTenant
    {
        public string Name { get => this.GetType().FullName; }

        public int Sequence { get => 0; }

        /// <summary>
        /// This method allows to check if a template can be provisioned in the currently selected target
        /// </summary>
        /// <param name="tenant">The target Tenant</param>
        /// <param name="hierarchy">The Template to hierarchy</param>
        /// <param name="sequenceId">The sequence to test within the hierarchy</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        public virtual CanProvisionResult CanProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // By default everything can be provisioned
            return (new CanProvisionResult { CanProvision = true, Issues = null });
        }

        protected CanProvisionResult EvaluateSiteRule<CanProvisionRuleSite>(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
            where CanProvisionRuleSite: CanProvisionRuleSiteBase
        {
            // Prepare the default output
            var result = new CanProvisionResult();

            // Target the root site collection
            string tenantRootsiteUrl = null;
#if !ONPREMISES
            tenant.EnsureProperty(t => t.RootSiteUrl);
            tenantRootsiteUrl = tenant.RootSiteUrl;
#else
            tenantRootsiteUrl = tenant.GetTenantRootSiteUrl();
            if (string.IsNullOrEmpty(tenantRootsiteUrl))
            {
                return result;
            }
#endif

            // Connect to the root site collection
            using (var context = tenant.Context.Clone(tenantRootsiteUrl))
            {
                // Evaluate the corresponding Site rule
                var innerRule = Activator.CreateInstance(typeof(CanProvisionRuleSite)) as CanProvisionRuleSiteBase;
                innerRule.TenantAdminSiteUrl = tenant.Context.Url;

                Model.ProvisioningTemplate dummyTemplate = null;

                // If we don't have templates
                if (hierarchy.Templates.Count == 0)
                {
                    dummyTemplate = new Model.ProvisioningTemplate();
                    dummyTemplate.Id = $"DUMMY-{Guid.NewGuid()}";
                    hierarchy.Templates.Add(dummyTemplate);
                }

                // Invoke the Site level rule
                result = innerRule.CanProvision(context.Web, hierarchy.Templates[0], applyingInformation);

                if (dummyTemplate != null)
                {
                    // Remove the dummy template, if any
                    hierarchy.Templates.Remove(dummyTemplate);
                }
            }

            return (result);

        }
    }
}
