using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class NavigationImplementation : ImplementationBase
    {

        internal void SiteCollectionNavigation(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Publishing needs to be activated for this test and we need a termset to be present
                ImplementPrerequisites(cc);

                #region Managed_Structural
                // Retrieved target xml data is not matching with source xml navigation types so changing navigation settings to get correct data.
                ChangeNavigationSettings(cc, StandardNavigationSource.TaxonomyProvider, StandardNavigationSource.PortalProvider);

                // Explicitely clear out the base template for this test as otherwise we're not getting any results back
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web)
                {
                    BaseTemplate = null,
                    HandlersToProcess = Handlers.Navigation,
                };

                var result = TestProvisioningTemplate(cc, "navigation_add_1605.xml", Handlers.Navigation, null, ptci);
                NavigationValidator nv = new NavigationValidator();
                nv.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv.Validate(result.SourceTemplate.Navigation, result.TargetTemplate.Navigation, result.SourceTokenParser));
                #endregion

                #region Structural_Managed
                ChangeNavigationSettings(cc, StandardNavigationSource.PortalProvider, StandardNavigationSource.TaxonomyProvider);

                var result2 = TestProvisioningTemplate(cc, "navigation_add2_1605.xml", Handlers.Navigation, null, ptci);
                NavigationValidator nv2 = new NavigationValidator();
                nv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv2.Validate(result2.SourceTemplate.Navigation, result2.TargetTemplate.Navigation, result2.SourceTokenParser));
                #endregion
            }
        }

        internal void WebNavigation(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Publishing needs to be activated for this test and we need a termset to be present
                ImplementPrerequisites(cc);

                #region Managed_Structural
                // Retrieved target xml data is not matching with source xml navigation types so changing navigation settings to get correct data.
                ChangeNavigationSettings(cc, StandardNavigationSource.TaxonomyProvider, StandardNavigationSource.PortalProvider);

                // Explicitely clear out the base template for this test as otherwise we're not getting any results back
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web)
                {
                    BaseTemplate = null,
                    HandlersToProcess = Handlers.Navigation,
                };

                var result = TestProvisioningTemplate(cc, "navigation_add_1605.xml", Handlers.Navigation, null, ptci);
                NavigationValidator nv = new NavigationValidator();
                nv.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv.Validate(result.SourceTemplate.Navigation, result.TargetTemplate.Navigation, result.SourceTokenParser));
                #endregion

                #region Structural_Managed
                ChangeNavigationSettings(cc, StandardNavigationSource.PortalProvider, StandardNavigationSource.TaxonomyProvider);

                var result2 = TestProvisioningTemplate(cc, "navigation_add2_1605.xml", Handlers.Navigation, null, ptci);
                NavigationValidator nv2 = new NavigationValidator();
                nv2.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv2.Validate(result2.SourceTemplate.Navigation, result2.TargetTemplate.Navigation, result2.SourceTokenParser));
                #endregion

                #region Inherit_Structural
                ChangeNavigationSettings(cc, StandardNavigationSource.InheritFromParentWeb, StandardNavigationSource.PortalProvider);

                var result3 = TestProvisioningTemplate(cc, "navigation_add3_1605.xml", Handlers.Navigation, null, ptci);
                NavigationValidator nv3 = new NavigationValidator();
                nv3.SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
                Assert.IsTrue(nv3.Validate(result3.SourceTemplate.Navigation, result3.TargetTemplate.Navigation, result3.SourceTokenParser));
                #endregion
            }
        }

        #region Helper methods
        // Retrieved target xml data is not matching with source xml navigation types so changing navigation settings to get correct data.
        public void ChangeNavigationSettings(ClientContext cc, StandardNavigationSource gSource, StandardNavigationSource cSource)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(cc);
            taxonomySession.UpdateCache();
            cc.Load(taxonomySession, ts => ts.TermStores);
            cc.ExecuteQueryRetry();

            var navigationSettings = new WebNavigationSettings(cc, cc.Web);
            navigationSettings.GlobalNavigation.Source = gSource;
            navigationSettings.CurrentNavigation.Source = cSource;
            navigationSettings.Update(taxonomySession);

            try
            {
                cc.ExecuteQueryRetry();
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
                cc.ExecuteQueryRetry();
            }
        }

        private void ImplementPrerequisites(ClientContext cc)
        {
            TestProvisioningTemplate(cc, "navigation_add_prereq.xml", Handlers.TermGroups | Handlers.Features);
        }
        #endregion

    }
}