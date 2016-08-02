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
    /// Test cases for the provisioning engine search settings functionality
    /// </summary>
    [TestClass]
   public class SearchSettingTests : FunctionalTestBase
    {
        #region Construction
        public SearchSettingTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_98a9f94f-acf6-4940-aef9-e2f2dc0e3d45";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_98a9f94f-acf6-4940-aef9-e2f2dc0e3d45/sub";
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
        /// Site Search Settings Test
        /// </summary>
        [TestMethod]
        public void SiteCollection1605SearchSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.IncludeSearchConfiguration = true;
                ptci.HandlersToProcess = Handlers.SearchSettings;

                var result = TestProvisioningTemplate(cc, "searchsettings_site_1605_add.xml", Handlers.SearchSettings, null, ptci);
                SearchSettingValidator sv = new SearchSettingValidator();
                Assert.IsTrue(sv.Validate(result.SourceTemplate.SiteSearchSettings, result.TargetTemplate.SiteSearchSettings));
            }

        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Web Search Settings test
        /// </summary>
        [TestMethod]
        public void Web1605SearchSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.IncludeSearchConfiguration = true;
                ptci.HandlersToProcess = Handlers.SearchSettings;

                var result = TestProvisioningTemplate(cc, "searchsettings_web_1605_add.xml", Handlers.SearchSettings, null, ptci);
                SearchSettingValidator sv = new SearchSettingValidator();
                Assert.IsTrue(sv.Validate(result.SourceTemplate.WebSearchSettings, result.TargetTemplate.WebSearchSettings));
            }
        }
        #endregion
    }
}
