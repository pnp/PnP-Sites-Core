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
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-03-FullSample-01.xml");

            CanProvisionResult result = null;

            using (var pnpContext = new PnPProvisioningContext())
            {
                using (var context = TestCommon.CreateClientContext())
                {
                    var applyingInformation = new ProvisioningTemplateApplyingInformation();
                    result = CanProvisionRulesManager.CanProvision(context.Web, hierarchy.Templates[0], applyingInformation);
                }
            }

            Assert.IsNotNull(result);
#if ONPREMISES
            // Because the "apps" rule is verified here
            Assert.IsFalse(result.CanProvision);
#else
            Assert.IsTrue(result.CanProvision);
#endif
        }

        [TestMethod]
        public void CanProvisionHierarchy()
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
                using (var tenantContext = TestCommon.CreateTenantClientContext())
                {
                    var applyingInformation = new ProvisioningTemplateApplyingInformation();
                    var tenant = new Tenant(tenantContext);
                    result = CanProvisionRulesManager.CanProvision(tenant, hierarchy, String.Empty, applyingInformation);
                }
            }

            Assert.IsNotNull(result);
            Assert.IsTrue(result.CanProvision);
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
