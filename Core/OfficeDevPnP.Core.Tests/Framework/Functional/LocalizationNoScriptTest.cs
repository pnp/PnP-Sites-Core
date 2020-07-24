using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !SP2013 && !SP2016 
    /// <summary>
    /// Test cases for the provisioning engine Publishing functionality
    /// </summary>
    [TestClass]
    public class LocalizationNoScriptTest : FunctionalTestBase
    {
        #region Construction
        public LocalizationNoScriptTest()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_5cef4b69-58d0-41d4-9ea6-06de3004b30f";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_5cef4b69-58d0-41d4-9ea6-06de3004b30f/sub";
        }
        #endregion

        #region Test setup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            ClassInitBase(context, true);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            ClassCleanupBase();
        }
        #endregion

        /// <summary>
        /// PnPLocalizationTest test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionsLocalizationTest()
        {
            new LocalizationImplementation().SiteCollectionsLocalization(centralSiteCollectionUrl);
        }
        /// <summary>
        /// PnPLocalizationTest test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebLocalizationTest()
        {
            new LocalizationImplementation().WebLocalization(centralSubSiteUrl);
        }
    }
#endif
}
