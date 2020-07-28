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
#if !SP2013 && !SP2016
    [TestClass]
    public class ListNoScriptTests : FunctionalTestBase
    {
        #region Construction
        public ListNoScriptTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c81e4b0d-0242-4c80-8272-18f13e759333";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c81e4b0d-0242-4c80-8272-18f13e759333/sub";
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
#endif
}
