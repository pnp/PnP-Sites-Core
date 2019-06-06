using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine feature functionality
    /// </summary>
    [TestClass]
    public class FeatureTests: FunctionalTestBase
    {
        #region Construction
        public FeatureTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_b456f6e7-d69d-4a19-abb1-fd7f8be33019";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_b456f6e7-d69d-4a19-abb1-fd7f8be33019/sub";
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
        /// Verifies if both the site and web scoped features are correctly activated/deactivated for a root web of a site collection
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionFeatureActivationDeactivationTest()
        {
            new FeatureImplementation().SiteCollectionFeatureActivationDeactivation(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Verifies is the web scoped features are correctly activated/deactivated for a sub site
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebFeatureActivationDeactivationTest()
        {
            new FeatureImplementation().WebFeatureActivationDeactivation(centralSubSiteUrl);
        }
        #endregion
    }
}
