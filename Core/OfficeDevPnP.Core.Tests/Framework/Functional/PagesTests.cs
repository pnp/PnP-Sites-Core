using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class PagesTests : FunctionalTestBase
    {
        #region Construction
        public PagesTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_06602218-e4fe-469a-8b51-95c6f491718e";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_06602218-e4fe-469a-8b51-95c6f491718e/sub";
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
        /// PagesTest Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionPagesTest()
        {
            new PagesImplementation().SiteCollectionPages(centralSiteCollectionUrl);
        }

        /// <summary>
        /// Client side pages Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionClientSidePagesTest()
        {
            new PagesImplementation().SiteCollectionClientSidePages(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// PagesTest Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebPagesTest()
        {
            new PagesImplementation().WebPages(centralSubSiteUrl);
        }

        /// <summary>
        /// Client side pages Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebClientSidePagesTest()
        {
            new PagesImplementation().WebClientSidePages(centralSubSiteUrl);
        }
        #endregion
    }
}
