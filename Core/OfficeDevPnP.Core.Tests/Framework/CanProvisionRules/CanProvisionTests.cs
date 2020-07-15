using System;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace OfficeDevPnP.Core.Tests.Framework.CanProvisionRules
{
    [TestClass]
    public class CanProvisionTests
    {
        [TestMethod]
        public void CanProvisionSite()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Template requires term store work, so this will not work in app-only");

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-03-FullSample-01.xml");

            var applyingInformation = new ProvisioningTemplateApplyingInformation();
            var template = hierarchy.Templates[0];
            if (TestCommon.AppOnlyTesting())
            {
                if (applyingInformation.HandlersToProcess.Has(Core.Framework.Provisioning.Model.Handlers.TermGroups)
                    || applyingInformation.HandlersToProcess.Has(Core.Framework.Provisioning.Model.Handlers.SearchSettings))
                {
                    bool templateSupportsAppOnly = this.IsTemplateSupportedForAppOnly(template);
                    if (!templateSupportsAppOnly)
                    {
                        Assert.Inconclusive("Taxonomy and SearchSettings tests are not supported when testing using app-only context.");
                    }
                }
            }

            CanProvisionResult result = null;

            using (var pnpContext = new PnPProvisioningContext())
            {
                using (var context = TestCommon.CreateClientContext())
                {
                    result = CanProvisionRulesManager.CanProvision(context.Web, hierarchy.Templates[0], applyingInformation);
                }
            }

            Assert.IsNotNull(result);
#if SP2013 || SP2016
            // Because the "apps" rule is verified here
            Assert.IsFalse(result.CanProvision);
#else
            Assert.IsTrue(result.CanProvision);
#endif
        }

        private bool IsTemplateSupportedForAppOnly(Core.Framework.Provisioning.Model.ProvisioningTemplate template)
        {
            bool result = true;

            if (template.TermGroups != null
                && template.TermGroups.Count > 0)
            {
                result = false;
            }
            else if (!string.IsNullOrEmpty(template.SiteSearchSettings))
            {
                result = false;
            }
            else if (!string.IsNullOrEmpty(template.WebSearchSettings))
            {
                result = false;
            }

            return result;
        }



        [TestMethod]
        public void CanProvisionHierarchy()
        {
            if (TestCommon.AppOnlyTesting()) Assert.Inconclusive("Template requires term store work, so this will not work in app-only");

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-03-FullSample-01.xml");

            var applyingInformation = new ProvisioningTemplateApplyingInformation();
            if (TestCommon.AppOnlyTesting())
            {
                bool templateSupportsAppOnly = true;

                if (applyingInformation.HandlersToProcess.Has(Core.Framework.Provisioning.Model.Handlers.TermGroups)
                    || applyingInformation.HandlersToProcess.Has(Core.Framework.Provisioning.Model.Handlers.SearchSettings))
                {
                    if (hierarchy.Templates.Count > 0)
                    {
                        foreach (var template in hierarchy.Templates)
                        {
                            templateSupportsAppOnly = this.IsTemplateSupportedForAppOnly(template);
                            if (!templateSupportsAppOnly)
                            {
                                break;
                            }
                        }
                    }
                }

                if (!templateSupportsAppOnly)
                {
                    Assert.Inconclusive("Taxonomy and SearchSettings tests are not supported when testing using app-only context.");
                }
            }

            CanProvisionResult result = null;

            using (var pnpContext = new PnPProvisioningContext())
            {
                using (var tenantContext = TestCommon.CreateTenantClientContext())
                {
                    var tenant = new Tenant(tenantContext);
                    result = CanProvisionRulesManager.CanProvision(tenant, hierarchy, String.Empty, applyingInformation);
                }
            }

            Assert.IsNotNull(result);
#if SP2013 || SP2016
            // Because the "apps" rule is verified here
            Assert.IsFalse(result.CanProvision);
#else
            Assert.IsTrue(result.CanProvision);
            Assert.IsTrue(result.CanProvision);
#endif
        }

        [TestMethod]
        public void CanProvisionOffice365()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-03-FullSample-01.xml");

            CanProvisionResult result = null;

            using (var pnpContext = new PnPProvisioningContext())
            {
                var applyingInformation = new ProvisioningTemplateApplyingInformation();
                result = CanProvisionRulesManager.CanProvision(hierarchy, String.Empty, applyingInformation);
            }

            Assert.IsNotNull(result);
            Assert.IsTrue(result.CanProvision);
        }
    }
}
