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
   public class RegionalSettingsTests : FunctionalTestBase
    {
        #region Construction
        public RegionalSettingsTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_23939b71-a69c-4d57-a818-5448a1cdaf28";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_23939b71-a69c-4d57-a818-5448a1cdaf28/sub";
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
        /// Site RegionalSettings Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionRegionalSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var result = TestProvisioningTemplate(cc, "regionalsettings_add.xml", Handlers.RegionalSettings);
                RegionalSettingsValidator rv = new RegionalSettingsValidator();
                Assert.IsTrue(rv.Validate(result.SourceTemplate.RegionalSettings, result.TargetTemplate.RegionalSettings, result.TargetTokenParser));

                var result2 = TestProvisioningTemplate(cc, "regionalsettings_delta_1.xml", Handlers.RegionalSettings);
                RegionalSettingsValidator rv2 = new RegionalSettingsValidator();
                Assert.IsTrue(rv2.Validate(result2.SourceTemplate.RegionalSettings, result2.TargetTemplate.RegionalSettings, result2.TargetTokenParser));
            }
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Web RegionalSettings test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebRegionalSettingsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                var result = TestProvisioningTemplate(cc, "regionalsettings_add.xml", Handlers.RegionalSettings);
                RegionalSettingsValidator rv = new RegionalSettingsValidator();
                Assert.IsTrue(rv.Validate(result.SourceTemplate.RegionalSettings, result.TargetTemplate.RegionalSettings, result.TargetTokenParser));

                var result2 = TestProvisioningTemplate(cc, "regionalsettings_delta_1.xml", Handlers.RegionalSettings);
                RegionalSettingsValidator rv2 = new RegionalSettingsValidator();
                Assert.IsTrue(rv2.Validate(result2.SourceTemplate.RegionalSettings, result2.TargetTemplate.RegionalSettings, result2.TargetTokenParser));
            }
        }
        #endregion
    }
}
