using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
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
   public class SearchsettingTest : FunctionalTestBase
    {
        #region Construction
        public SearchsettingTest()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://crtlab2.sharepoint.com/sites/source2";
            //centralSubSiteUrl = "https://crtlab2.sharepoint.com/sites/source2/sub";
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

        #region SearchsettingTest
        /// <summary>
        /// Site Search Settings Test
        /// </summary>
        [TestMethod]
        public void SiteCollectionSearchSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var result = TestProvisioningTemplate(cc, "searchsettings_site_add .xml", Handlers.SearchSettings);
                SearchsettingsValidator searchVal = new SearchsettingsValidator();
                string targetSearchSettings = result.TargetTemplate.SiteSearchSettings;
                if (targetSearchSettings == null)
                {
                    targetSearchSettings = result.TargetTemplate.SearchSettings;
                }

                Assert.IsTrue(searchVal.Validate(result.SourceTemplate.SiteSearchSettings, targetSearchSettings));
            }

        }
        /// <summary>
        /// Web Seacrh Settings test
        /// </summary>
        [TestMethod]
        public void WebSearchSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                var result = TestProvisioningTemplate(cc, "searchsettings_web_add.xml", Handlers.SearchSettings);
                SearchsettingsValidator searchVal = new SearchsettingsValidator();
                string targetSearchSettings = result.TargetTemplate.WebSearchSettings;
                if (targetSearchSettings == null)
                {
                    targetSearchSettings = result.TargetTemplate.SearchSettings;
                }

                Assert.IsTrue(searchVal.Validate(result.SourceTemplate.WebSearchSettings, targetSearchSettings));
            }

        }
        #endregion
    }
}
