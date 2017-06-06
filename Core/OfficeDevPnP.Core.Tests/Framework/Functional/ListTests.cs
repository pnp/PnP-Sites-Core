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
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_811105dc-86c9-4b27-bf8d-6f275f2176e8";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_811105dc-86c9-4b27-bf8d-6f275f2176e8/sub";
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
        #endregion

    }
}
