using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
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
    /// Summary description for NavigationTests
    /// </summary>
    [TestClass]
    public class NavigationTests : FunctionalTestBase
    {
        #region construction
        public NavigationTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_f5f12c31-8aae-43f1-81d6-c389cb1c7505";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_f5f12c31-8aae-43f1-81d6-c389cb1c7505/sub";
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
        /// Navigation test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionNavigationTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                // Publishing needs to be activated for this test
                Prerequisite_EnablePublishingFeatures(cc);

                // Termset is required to choose navigation term in managed navigation section
                Prerequisite_CreateTermGroup(cc);

                #region Managed_Structural
                // Retrieved target xml data is not matching with source xml navigation types so changing navigation settings to get correct data.
                ChangeNavigationSettings(cc, StandardNavigationSource.TaxonomyProvider, StandardNavigationSource.PortalProvider);

                var result = TestProvisioningTemplate(cc, "navigation_add_1605.xml", Handlers.Navigation);                
                NavigationValidator nv = new NavigationValidator();
                nv.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv.Validate(result.SourceTemplate.Navigation, result.TargetTemplate.Navigation, result.SourceTokenParser));

                #endregion

                #region Structural_Managed
                ChangeNavigationSettings(cc, StandardNavigationSource.PortalProvider, StandardNavigationSource.TaxonomyProvider);

                var result2 = TestProvisioningTemplate(cc, "navigation_add2_1605.xml", Handlers.Navigation);
                NavigationValidator nv2 = new NavigationValidator();
                nv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv2.Validate(result2.SourceTemplate.Navigation, result2.TargetTemplate.Navigation, result2.SourceTokenParser));
                #endregion
            }
        }

        #endregion

        #region WebTest
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebNavigationTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                // Publishing needs to be activated for this test
                Prerequisite_EnablePublishingFeatures(cc);

                #region Managed_Structural
                // Retrieved target xml data is not matching with source xml navigation types so changing navigation settings to get correct data.
                ChangeNavigationSettings(cc, StandardNavigationSource.TaxonomyProvider, StandardNavigationSource.PortalProvider);

                var result = TestProvisioningTemplate(cc, "navigation_add_1605.xml", Handlers.Navigation);
                NavigationValidator nv = new NavigationValidator();
                nv.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv.Validate(result.SourceTemplate.Navigation, result.TargetTemplate.Navigation, result.SourceTokenParser));

                #endregion

                #region Structural_Managed
                ChangeNavigationSettings(cc, StandardNavigationSource.PortalProvider, StandardNavigationSource.TaxonomyProvider);

                var result2 = TestProvisioningTemplate(cc, "navigation_add2_1605.xml", Handlers.Navigation);
                NavigationValidator nv2 = new NavigationValidator();
                nv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv2.Validate(result2.SourceTemplate.Navigation, result2.TargetTemplate.Navigation, result2.SourceTokenParser));
                #endregion

                #region Inherit_Structural
                ChangeNavigationSettings(cc, StandardNavigationSource.InheritFromParentWeb, StandardNavigationSource.PortalProvider);

                var result3 = TestProvisioningTemplate(cc, "navigation_add3_1605.xml", Handlers.Navigation);
                NavigationValidator nv3 = new NavigationValidator();
                nv3.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv3.Validate(result3.SourceTemplate.Navigation, result3.TargetTemplate.Navigation, result3.SourceTokenParser));
                #endregion
            }
        }
        #endregion

        #region Helper methods

        // Retrieved target xml data is not matching with source xml navigation types so changing navigation settings to get correct data.
        public void ChangeNavigationSettings(ClientContext cc, StandardNavigationSource gSource, StandardNavigationSource cSource)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(cc);
            taxonomySession.UpdateCache();
            cc.Load(taxonomySession, ts => ts.TermStores);
            cc.ExecuteQuery();

            var navigationSettings = new WebNavigationSettings(cc, cc.Web);
            navigationSettings.GlobalNavigation.Source = gSource;
            navigationSettings.CurrentNavigation.Source = cSource;
            navigationSettings.Update(taxonomySession);

            try
            {
                cc.ExecuteQuery();                
            }
            catch (Exception ex) // if termset not found then set newly created termset to managed navigation
            {
                TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                Microsoft.SharePoint.Client.Taxonomy.TermGroup group = termStore.GetTermGroupByName("TG_1"); // TG_1 is a term group mentioned in navigation_add_1605.xml
                Microsoft.SharePoint.Client.Taxonomy.TermSet termset = group.TermSets.GetByName("TS_1_1"); // TS_1_1 is a term set mentioned in navigation_add_1605.xml
                cc.Load(termStore);
                cc.Load(group, g => g.TermSets);
                cc.Load(termset);
                cc.ExecuteQuery();

                if (StandardNavigationSource.TaxonomyProvider == gSource)
                {
                    navigationSettings.GlobalNavigation.TermStoreId = termStore.Id;
                    navigationSettings.GlobalNavigation.TermSetId = termset.Id;
                }

                if (StandardNavigationSource.TaxonomyProvider == cSource)
                {
                    navigationSettings.CurrentNavigation.TermStoreId = termStore.Id;
                    navigationSettings.CurrentNavigation.TermSetId = termset.Id;
                }

                navigationSettings.GlobalNavigation.Source = gSource;
                navigationSettings.CurrentNavigation.Source = cSource;
                navigationSettings.Update(taxonomySession);
                cc.ExecuteQuery();
            }
        }

        private void Prerequisite_CreateTermGroup(ClientContext cc)
        {
            TestProvisioningTemplate(cc, "navigation_add_1605.xml", Handlers.TermGroups);
        }

        private void Prerequisite_EnablePublishingFeatures(ClientContext cc)
        {
            TestProvisioningTemplate(cc, "navigation_add_prereq.xml", Handlers.Features);
        }
        #endregion

    }
}
