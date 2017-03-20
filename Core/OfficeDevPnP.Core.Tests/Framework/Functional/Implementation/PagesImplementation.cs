using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class PagesImplementation : ImplementationBase
    {
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
    }
}