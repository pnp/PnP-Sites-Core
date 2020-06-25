using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class TermGroupImplementation : ImplementationBase
    {
        internal void SiteCollectionTermGroup(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // delete termgroups first
                DeleteTermGroups(cc);

                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.IncludeAllTermGroups = true;
                ptci.IncludeSiteCollectionTermGroup = true;
                ptci.IncludeTermGroupsSecurity = true;
                ptci.HandlersToProcess = Handlers.TermGroups;

                var result = TestProvisioningTemplate(cc, "termgroup_add.xml", Handlers.TermGroups, null, ptci);
                TermGroupValidator tv = new TermGroupValidator();
                Assert.IsTrue(tv.Validate(result.SourceTemplate.TermGroups, result.TargetTemplate.TermGroups, result.TargetTokenParser));

                var result2 = TestProvisioningTemplate(cc, "termgroup_delta_1605.xml", Handlers.TermGroups, null, ptci);
                TermGroupValidator tv2 = new TermGroupValidator();
                tv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(tv2.Validate(result2.SourceTemplate.TermGroups, result2.TargetTemplate.TermGroups, result2.TargetTokenParser));

#if !ONPREMISES
                var result3 = TestProvisioningTemplate(cc, "termgroup_delta_1605_1.xml", Handlers.TermGroups, null, ptci);
                TermGroupValidator tv3 = new TermGroupValidator();
                tv3.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(tv3.Validate(result3.SourceTemplate.TermGroups, result3.TargetTemplate.TermGroups, result3.TargetTokenParser));
#endif
            }
        }


        internal void WebTermGroup(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // delete termgroups first
                DeleteTermGroups(cc);

                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.IncludeAllTermGroups = true;
                ptci.IncludeSiteCollectionTermGroup = true;
                ptci.IncludeTermGroupsSecurity = true;
                ptci.HandlersToProcess = Handlers.TermGroups;

                var result = TestProvisioningTemplate(cc, "termgroup_add.xml", Handlers.TermGroups, null, ptci);
                TermGroupValidator tv = new TermGroupValidator();
                Assert.IsTrue(tv.Validate(result.SourceTemplate.TermGroups, result.TargetTemplate.TermGroups, result.TargetTokenParser));

                var result2 = TestProvisioningTemplate(cc, "termgroup_delta_1605.xml", Handlers.TermGroups, null, ptci);
                TermGroupValidator tv2 = new TermGroupValidator();
                tv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(tv2.Validate(result2.SourceTemplate.TermGroups, result2.TargetTemplate.TermGroups, result2.TargetTokenParser));

#if !ONPREMISES
                var result3 = TestProvisioningTemplate(cc, "termgroup_delta_1605_1.xml", Handlers.TermGroups, null, ptci);
                TermGroupValidator tv3 = new TermGroupValidator();
                tv3.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(tv3.Validate(result3.SourceTemplate.TermGroups, result3.TargetTemplate.TermGroups, result3.TargetTokenParser));
#endif
            }
        }


        #region Helper methods
        private void DeleteTermGroups(ClientContext cc)
        {
            var taxSession = TaxonomySession.GetTaxonomySession(cc);
            var termStore = taxSession.GetDefaultSiteCollectionTermStore();



            // Ensure that the group is empty before deleting it. 
            // exceptions like the following happen:
            // Microsoft.SharePoint.Client.ServerException: Microsoft.SharePoint.Client.ServerException: A Group cannot be deleted unless it is empty..

            OfficeDevPnP.Core.Tests.Utilities.RetryHelper.Do(
                () => this.DeleteTermGroupsImplementation(cc, termStore),
                TimeSpan.FromSeconds(10),
                3);
        }

        private void DeleteTermGroupsImplementation(ClientContext cc, TermStore termStore)
        {
            cc.Load(termStore.Groups, 
                p => p.Include(
                    t => t.Name, 
                    t => t.TermSets.Include(
                        s => s.Name,
                        s => s.Terms.Include(
                            q => q.IsDeprecated, 
                            q => q.ReusedTerms
                            )
                        )
                    )
                );

            cc.ExecuteQueryRetry();

            foreach (var termGroup in termStore.Groups.ToList())
            {
                DeleteTermGroupImplementation(cc, termGroup);
            }

            var siteCollectionGroup = termStore.GetSiteCollectionGroup(cc.Site, true);
            cc.Load(siteCollectionGroup, 
                t => t.Name, 
                t => t.TermSets.Include(
                    s => s.Name, 
                    s => s.Terms.Include(
                        q => q.IsDeprecated, 
                        q => q.ReusedTerms
                        )
                    )
                );
            cc.ExecuteQueryRetry();
            DeleteTermGroupImplementation(cc, siteCollectionGroup, true);

            termStore.CommitAll();
            termStore.UpdateCache();
            cc.ExecuteQueryRetry();
        }

        private static void DeleteTermGroupImplementation(ClientContext cc, Microsoft.SharePoint.Client.Taxonomy.TermGroup termGroup, bool siteCollectionGroup = false)
        {
            if (termGroup.Name.StartsWith("TG_") || siteCollectionGroup)
            {
                foreach (var termSet in termGroup.TermSets)
                {
                    if (termSet.Name.StartsWith("TS_"))
                    {
                        foreach (var term in termSet.Terms)
                        {
                            // first deleted the reused terms to avoid issues with re-using the same term id in an upcoming test run
                            foreach (var reusedTerm in term.ReusedTerms)
                            {
                                term.DeleteObject();
                            }
                            cc.ExecuteQueryRetry();
                        }

                        termSet.DeleteObject();
                    }
                }

                if (!siteCollectionGroup)
                {
                    termGroup.DeleteObject();
                }
                cc.ExecuteQueryRetry();
            }
        }
        #endregion

    }
}