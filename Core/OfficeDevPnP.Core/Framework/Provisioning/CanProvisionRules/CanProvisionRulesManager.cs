using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// This class manages all the CanProvision rules
    /// </summary>
    public static class CanProvisionRulesManager
    {
        /// <summary>
        /// This method allows to check if a template can be provisioned in the currently selected target
        /// </summary>
        /// <param name="web">The target Web</param>
        /// <param name="template">The Template to provision</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        public static CanProvisionResult CanProvision(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            CanProvisionResult result = new CanProvisionResult();

            List<ICanProvisionRuleSite> rules = GetCanProvisionRules<ICanProvisionRuleSite>();

            foreach (var rule in rules)
            {
                var ruleResult = rule.CanProvision(web, template, applyingInformation);
                result.CanProvision &= ruleResult.CanProvision;
                result.Issues.AddRange(ruleResult.Issues);
            }

            return (result);
        }

        /// <summary>
        /// This method allows to check if a template can be provisioned in the currently selected target
        /// </summary>
        /// <param name="tenant">The target Tenant</param>
        /// <param name="hierarchy">The Template to hierarchy</param>
        /// <param name="sequenceId">The sequence to test within the hierarchy</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        public static CanProvisionResult CanProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            CanProvisionResult result = new CanProvisionResult();

            List<ICanProvisionRuleTenant> rules = GetCanProvisionRules<ICanProvisionRuleTenant>();

            foreach (var rule in rules)
            {
                var ruleResult = rule.CanProvision(tenant, hierarchy, sequenceId, applyingInformation);
                result.CanProvision &= ruleResult.CanProvision;
                result.Issues.AddRange(ruleResult.Issues);
            }

            return (result);
        }

        /// <summary>
        /// This method allows to check if a template can be provisioned
        /// </summary>
        /// <param name="hierarchy">The Template to hierarchy</param>
        /// <param name="sequenceId">The sequence to test within the hierarchy</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        public static CanProvisionResult CanProvision(Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            CanProvisionResult result = new CanProvisionResult();

            List<ICanProvisionRuleOffice365> rules = GetCanProvisionRules<ICanProvisionRuleOffice365>();

            foreach (var rule in rules)
            {
                var ruleResult = rule.CanProvision(hierarchy, sequenceId, applyingInformation);
                result.CanProvision &= ruleResult.CanProvision;
                result.Issues.AddRange(ruleResult.Issues);
            }

            return (result);
        }

        private static List<TCanProvisionRule> GetCanProvisionRules<TCanProvisionRule>()
            where TCanProvisionRule : ICanProvisionRuleBase
        {
            // Get all the rules to run in automated mode, ordered by Sequence
            var currentAssembly = typeof(CanProvisionRulesManager).Assembly;

            // Get all the rules for the target
            var ruleTypes = currentAssembly.GetTypes()
                .Where(t =>
                {
                    var a = t.GetCustomAttributes<CanProvisionRuleAttribute>(false).FirstOrDefault();
                    if (a != null)
                    {
                        if (typeof(TCanProvisionRule) == typeof(ICanProvisionRuleSite))
                        {
                            return (a.Scope == CanProvisionScope.Site);
                        }
                        else if (typeof(TCanProvisionRule) == typeof(ICanProvisionRuleTenant))
                        {
                            return (a.Scope == CanProvisionScope.Tenant);
                        }
                        else if (typeof(TCanProvisionRule) == typeof(ICanProvisionRuleOffice365))
                        {
                            return (a.Scope == CanProvisionScope.Office365);
                        }
                        else
                        {
                            return (false);
                        }
                    }
                    else
                    {
                        return (false);
                    }
                })
                .OrderBy(t =>
                {
                    var a = t.GetCustomAttributes<CanProvisionRuleAttribute>(false).FirstOrDefault();
                    return (a.Sequence);
                });
            ;


            var rules = new List<TCanProvisionRule>();

            foreach (var ruleType in ruleTypes)
            {
                var rule = (TCanProvisionRule)Activator.CreateInstance(ruleType);
                rules.Add(rule);
            }

            return rules;
        }
    }
}
