using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class ListTests : FunctionalTestBase
    {
        #region Construction
        public ListTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_2f75da63-39c4-456f-af8c-14adf41074f6";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_2f75da63-39c4-456f-af8c-14adf41074f6/sub";
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
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionListAddingTest()
        {
            new ListImplementation().SiteCollectionListAdding(centralSiteCollectionUrl);
        }

        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollection1605ListAddingTest()
        {
            new ListImplementation().SiteCollection1605ListAdding(centralSiteCollectionUrl);
        }

        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollection1705ListAddingTest()
        {
            new ListImplementation().SiteCollection1705ListAdding(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebListAddingTest()
        {
            new ListImplementation().WebListAdding(centralSiteCollectionUrl, centralSubSiteUrl);
        }

        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void Web1605ListAddingTest()
        {
            new ListImplementation().Web1605ListAdding(centralSubSiteUrl);
        }

        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void Web1705ListAddingTest()
        {
            new ListImplementation().Web1705ListAdding(centralSubSiteUrl);
        }
        #endregion

    }
}
