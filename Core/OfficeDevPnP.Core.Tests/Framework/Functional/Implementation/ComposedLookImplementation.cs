using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class ComposedLookImplementation : ImplementationBase
    {
        internal void SiteCollectionComposedLook(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                if (!cc.Web.IsNoScriptSite())
                {
                    // Add supporting files
                    TestProvisioningTemplate(cc, "composedlook_files.xml", Handlers.Files);

                    var result = TestProvisioningTemplate(cc, "composedlook_add_1.xml", Handlers.ComposedLook);
                    ComposedLookValidator composedLookVal = new ComposedLookValidator();
                    Assert.IsTrue(composedLookVal.Validate(result.SourceTemplate.ComposedLook, result.TargetTemplate.ComposedLook));

                    var result2 = TestProvisioningTemplate(cc, "composedlook_add_2.xml", Handlers.ComposedLook);
                    Assert.IsTrue(composedLookVal.Validate(result2.SourceTemplate.ComposedLook, result2.TargetTemplate.ComposedLook));
                }
            }
        }

        public void WebComposedLook(string siteCollectionUrl, string url)
        {
            using (var cc = TestCommon.CreateClientContext(siteCollectionUrl))
            {
                // Add supporting files
                TestProvisioningTemplate(cc, "composedlook_files.xml", Handlers.Files);
            }

            using (var cc = TestCommon.CreateClientContext(url))
            {
                if (!cc.Web.IsNoScriptSite())
                {
                    var result = TestProvisioningTemplate(cc, "composedlook_add_1.xml", Handlers.ComposedLook);
                    ComposedLookValidator composedLookVal = new ComposedLookValidator();
                    Assert.IsTrue(composedLookVal.Validate(result.SourceTemplate.ComposedLook, result.TargetTemplate.ComposedLook));

                    var result2 = TestProvisioningTemplate(cc, "composedlook_add_2.xml", Handlers.ComposedLook);
                    Assert.IsTrue(composedLookVal.Validate(result2.SourceTemplate.ComposedLook, result2.TargetTemplate.ComposedLook));
                }
            }
        }
    }
}