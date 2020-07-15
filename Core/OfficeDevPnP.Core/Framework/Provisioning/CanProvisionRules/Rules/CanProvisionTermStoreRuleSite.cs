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
    [CanProvisionRule(Scope = CanProvisionScope.Site, Sequence = 200)]
    internal class CanProvisionTermStoreRuleSite : CanProvisionRuleSiteBase
    {
        public override CanProvisionResult CanProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {

            // Prepare the default output
            var result = new CanProvisionResult();
#if !SP2013 && !SP2016

            Model.ProvisioningTemplate targetTemplate = null;

            if (template.ParentHierarchy != null)
            {
                // If we have a hierarchy, search for a template with Taxonomy settings, if any
                targetTemplate = template.ParentHierarchy.Templates.FirstOrDefault(t => t.TermGroups.Count > 0);

                if (targetTemplate == null)
                {
                    // or use the first in the hierarchy
                    targetTemplate = template.ParentHierarchy.Templates[0];
                }
            }
            else
            {
                // Otherwise, use the provided template
                targetTemplate = template;
            }


            // Verify if we need the Term Store permissions (i.e. the template contains term groups to provision, or sequences with TermStore settings)
            if ((targetTemplate.TermGroups != null && targetTemplate.TermGroups?.Count > 0) ||
                targetTemplate.ParentHierarchy.Sequences.Any(
                    s => s.TermStore?.TermGroups != null && s.TermStore?.TermGroups?.Count > 0))
            {
                using (var scope = new PnPMonitoredScope(this.Name))
                {
                    try
                    {
                        // Try to access the Term Store
                        TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(web.Context);
                        TermStore termStore = taxSession.GetDefaultKeywordsTermStore();
                        web.Context.Load(termStore,
                            ts => ts.Languages,
                            ts => ts.DefaultLanguage,
                            ts => ts.Groups.Include(
                                tg => tg.Name,
                                tg => tg.Id,
                                tg => tg.TermSets.Include(
                                    tset => tset.Name,
                                    tset => tset.Id)));
                        var siteCollectionTermGroup = termStore.GetSiteCollectionGroup((web.Context as ClientContext).Site, false);
                        web.Context.Load(siteCollectionTermGroup);
                        web.Context.ExecuteQueryRetry();

                        var termGroupId = Guid.NewGuid();
                        var group = termStore.CreateGroup($"Temp-{termGroupId.ToString()}", termGroupId);
                        termStore.CommitAll();
                        web.Context.Load(group);
                        web.Context.ExecuteQueryRetry();

                        // Now delete the just created termGroup, to cleanup the Term Store
                        group.DeleteObject();
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        // And if we fail, raise a CanProvisionIssue
                        result.CanProvision = false;
                        result.Issues.Add(new CanProvisionIssue()
                        {
                            Source = this.Name,
                            Tag = CanProvisionIssueTags.MISSING_TERMSTORE_PERMISSIONS,
                            Message = CanProvisionIssuesMessages.Term_Store_Not_Admin,
                            ExceptionMessage = ex.Message, // Here we have a specific exception
                            ExceptionStackTrace = ex.StackTrace, // Here we have a specific exception
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
