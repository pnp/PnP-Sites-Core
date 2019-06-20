using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class WebSettingsImplementation : ImplementationBase
    {
        internal void SiteCollectionWebSettings(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                // Add supporting files
                TestProvisioningTemplate(cc, "websettings_files.xml", Handlers.Files);

                var result = TestProvisioningTemplate(cc, "websettings_add.xml", Handlers.WebSettings);
                WebSettingsValidator wv = new WebSettingsValidator(cc);
                Assert.IsTrue(wv.Validate(result.SourceTemplate.WebSettings, result.TargetTemplate.WebSettings, result.TargetTokenParser));
            }
        }

        internal void SiteCollectionAuditSettings(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "auditsettings_add.xml", Handlers.AuditSettings);
                AuditSettingsValidator av = new AuditSettingsValidator(cc);
                Assert.IsTrue(av.Validate(result.SourceTemplate.AuditSettings, result.TargetTemplate.AuditSettings, result.TargetTokenParser));
            }
        }

        internal void WebWebSettings(string siteCollectionUrl, string url)
        {
            using (var cc = TestCommon.CreateClientContext(siteCollectionUrl))
            {
                // Add supporting files
                TestProvisioningTemplate(cc, "websettings_files.xml", Handlers.Files);
            }

            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "websettings_add.xml", Handlers.WebSettings);
                WebSettingsValidator wv = new WebSettingsValidator(cc);
                Assert.IsTrue(wv.Validate(result.SourceTemplate.WebSettings, result.TargetTemplate.WebSettings, result.TargetTokenParser));
            }
        }
    }
}