using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !SP2013 && !SP2016
    [TestClass]
    public class CustomActionNoScriptTests: FunctionalTestBase
    {
    #region Construction
        public CustomActionNoScriptTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_ab5f2990-6015-48c5-a09b-685153dcebc9";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_ab5f2990-6015-48c5-a09b-685153dcebc9/sub";
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
        public void SiteCollectionCustomActionAddingTest()
        {
            new CustomActionImplementation().SiteCollectionCustomActionAdding(centralSiteCollectionUrl);
        }
    #endregion

    #region Web test cases
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebCustomActionAddingTest()
        {
            new CustomActionImplementation().WebCustomActionAdding(centralSubSiteUrl);
        }
    #endregion
    }
#endif
}
