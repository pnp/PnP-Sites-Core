using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !ONPREMISES
    /// <summary>
    /// Summary description for NavigationTests
    /// </summary>
    [TestClass]
    public class NavigationNoScriptTests : FunctionalTestBase
    {
        #region construction
        public NavigationNoScriptTests()
        {
            isNoScriptSite = true;
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_f5f12c31-8aae-43f1-81d6-c389cb1c7505";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_f5f12c31-8aae-43f1-81d6-c389cb1c7505/sub";
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

        [TestInitialize()]
        public override void Initialize()
        {
            base.Initialize();

            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Test that require taxonomy creation are not supported in app-only.");
            }
        }
        #endregion

        #region Site collection test cases
        /// <summary>
        /// Navigation test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionNavigationTest()
        {
            new NavigationImplementation().SiteCollectionNavigation(centralSiteCollectionUrl);
        }

        #endregion

        #region WebTest
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebNavigationTest()
        {
            new NavigationImplementation().SiteCollectionNavigation(centralSubSiteUrl);
        }
        #endregion
    }
#endif
}
