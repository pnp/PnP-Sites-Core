using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class LanguageImplementation : ImplementationBase
    {
        internal void SiteCollectionLanguageSettings(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "languagesettings_add.xml", Handlers.SupportedUILanguages);
                LanguageSettingsValidator lv = new LanguageSettingsValidator();
                Assert.IsTrue(lv.Validate(result.SourceTemplate.SupportedUILanguages, result.TargetTemplate.SupportedUILanguages, result.TargetTokenParser));

                // Delta test: check if we also can remove a set language
                var result2 = TestProvisioningTemplate(cc, "languagesettings_delta.xml", Handlers.SupportedUILanguages);
                Assert.IsTrue(lv.Validate(result2.SourceTemplate.SupportedUILanguages, result2.TargetTemplate.SupportedUILanguages, result2.TargetTokenParser));
            }
        }

        internal void WebLanguageSettings(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "languagesettings_add.xml", Handlers.SupportedUILanguages);
                LanguageSettingsValidator lv = new LanguageSettingsValidator();
                Assert.IsTrue(lv.Validate(result.SourceTemplate.SupportedUILanguages, result.TargetTemplate.SupportedUILanguages, result.TargetTokenParser));

                // Delta test: check if we also can remove a set language
                var result2 = TestProvisioningTemplate(cc, "languagesettings_delta.xml", Handlers.SupportedUILanguages);
                Assert.IsTrue(lv.Validate(result2.SourceTemplate.SupportedUILanguages, result2.TargetTemplate.SupportedUILanguages, result2.TargetTokenParser));
            }
        }
    }
}