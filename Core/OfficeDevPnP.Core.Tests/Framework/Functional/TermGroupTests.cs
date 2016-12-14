using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine term group functionality
    /// </summary>
    [TestClass]
   public class TermGroupTests : FunctionalTestBase
    {
        #region Construction
        public TermGroupTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_d644f1c6-80ac-4858-8e63-a7a5ce26c206";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_d644f1c6-80ac-4858-8e63-a7a5ce26c206/sub";
        }
        #endregion

        #region Test setup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            ClassInitBase(context);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            ClassCleanupBase();
        }

        [TestInitialize()]
        public override void Initialize()
        {
            base.Initialize();

            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Test that require taxonomy creation are not supported in app-only.");
            }
        }
        #endregion

        #region Site collection test cases
        /// <summary>
        /// Site TermGroup Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionTermGroupTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
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
            }
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Web TermGroup test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebTermGroupTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
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

            }
        }
        #endregion

        #region Helper methods
        private void DeleteTermGroups(ClientContext cc)
        {
            var taxSession = TaxonomySession.GetTaxonomySession(cc);
            var termStore = taxSession.GetDefaultSiteCollectionTermStore();

            DeleteTermGroupsImplementation(cc, termStore);
        }

        private void DeleteTermGroupsImplementation(ClientContext cc, TermStore termStore)
        {
            cc.Load(termStore.Groups, p => p.Include(t => t.Name, t => t.TermSets.Include(s => s.Name, s => s.Terms.Include(q => q.IsDeprecated, q => q.ReusedTerms))));
            cc.ExecuteQueryRetry();

            foreach (var termGroup in termStore.Groups.ToList())
            {
                DeleteTermGroupImplementation(cc, termGroup);
            }

            var siteCollectionGroup = termStore.GetSiteCollectionGroup(cc.Site, true);
            cc.Load(siteCollectionGroup, t => t.Name, t => t.TermSets.Include(s => s.Name, s => s.Terms.Include(q => q.IsDeprecated, q => q.ReusedTerms)));
            cc.ExecuteQueryRetry();
            DeleteTermGroupImplementation(cc, siteCollectionGroup, true);

            termStore.CommitAll();
            termStore.UpdateCache();
            cc.ExecuteQueryRetry();
        }

        private static void DeleteTermGroupImplementation(ClientContext cc, Microsoft.SharePoint.Client.Taxonomy.TermGroup termGroup, bool siteCollectionGroup=false)
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
