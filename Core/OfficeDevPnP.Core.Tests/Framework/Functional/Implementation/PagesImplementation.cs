using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class PagesImplementation : ImplementationBase
    {

        #region classic pages
        internal void SiteCollectionPages(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "pages_add.xml", Handlers.Pages);
                PagesValidator pv = new PagesValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.Pages, cc));
            }
        }

        internal void WebPages(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "pages_add.xml", Handlers.Pages);
                PagesValidator pv = new PagesValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.Pages, cc));
            }
        }
        #endregion

        #region client side pages
        internal void SiteCollectionClientSidePages(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "clientsidepages_add_1705.xml", Handlers.Pages);
                ClientSidePagesValidator pv = new ClientSidePagesValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.ClientSidePages, cc));
            }
        }

        internal void WebClientSidePages(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "clientsidepages_add_1705.xml", Handlers.Pages);
                ClientSidePagesValidator pv = new ClientSidePagesValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.ClientSidePages, cc));
            }
        }

        #endregion
    }
}