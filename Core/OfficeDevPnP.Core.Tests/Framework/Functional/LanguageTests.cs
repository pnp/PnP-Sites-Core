using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
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
    /// Test cases for the provisioning engine web settings functionality
    /// </summary>
    [TestClass]
   public class LanguageTests : FunctionalTestBase
    {
        #region Construction
        public LanguageTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_6963f04e-da9d-4551-a823-94482982f862";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_6963f04e-da9d-4551-a823-94482982f862/sub";
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
        #endregion

        #region Site collection test cases
        /// <summary>
        /// Site LanguageSettings Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionLanguageSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var result = TestProvisioningTemplate(cc, "languagesettings_add.xml", Handlers.SupportedUILanguages);
                LanguageSettingsValidator lv = new LanguageSettingsValidator();
                Assert.IsTrue(lv.Validate(result.SourceTemplate.SupportedUILanguages, result.TargetTemplate.SupportedUILanguages, result.TargetTokenParser));

                // Delta test: check if we also can remove a set language
                var result2 = TestProvisioningTemplate(cc, "languagesettings_delta.xml", Handlers.SupportedUILanguages);
                Assert.IsTrue(lv.Validate(result2.SourceTemplate.SupportedUILanguages, result2.TargetTemplate.SupportedUILanguages, result2.TargetTokenParser));
            }
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Web WebSettings test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebLanguageSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                var result = TestProvisioningTemplate(cc, "languagesettings_add.xml", Handlers.SupportedUILanguages);
                LanguageSettingsValidator lv = new LanguageSettingsValidator();
                Assert.IsTrue(lv.Validate(result.SourceTemplate.SupportedUILanguages, result.TargetTemplate.SupportedUILanguages, result.TargetTokenParser));

                // Delta test: check if we also can remove a set language
                var result2 = TestProvisioningTemplate(cc, "languagesettings_delta.xml", Handlers.SupportedUILanguages);
                Assert.IsTrue(lv.Validate(result2.SourceTemplate.SupportedUILanguages, result2.TargetTemplate.SupportedUILanguages, result2.TargetTokenParser));
            }
        }
        #endregion
    }
}
