using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class SearchSettingsImplementation : ImplementationBase
    {
        internal void SiteCollection1605SearchSettings(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.IncludeSearchConfiguration = true;
                ptci.HandlersToProcess = Handlers.SearchSettings;

                var result = TestProvisioningTemplate(cc, "searchsettings_site_1605_add.xml", Handlers.SearchSettings, null, ptci);
                SearchSettingValidator sv = new SearchSettingValidator();
                Assert.IsTrue(sv.Validate(result.SourceTemplate.SiteSearchSettings, result.TargetTemplate.SiteSearchSettings));
            }
        }

        internal void Web1605SearchSettings(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.IncludeSearchConfiguration = true;
                ptci.HandlersToProcess = Handlers.SearchSettings;

                var result = TestProvisioningTemplate(cc, "searchsettings_web_1605_add.xml", Handlers.SearchSettings, null, ptci);
                SearchSettingValidator sv = new SearchSettingValidator();
                Assert.IsTrue(sv.Validate(result.SourceTemplate.WebSearchSettings, result.TargetTemplate.WebSearchSettings));
            }
        }

    }
}