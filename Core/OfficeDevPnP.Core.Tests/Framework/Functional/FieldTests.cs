using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class FieldTests : FunctionalTestBase
    {
        #region Construction
        public FieldTests()
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
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionFieldAddingTest()
        {
            new FieldImplementation().SiteCollectionFieldAdding(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        // No need to have these as the engine is blocking creation and extraction of fields at web level
        #endregion


    }
}
