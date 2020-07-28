using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Test cases for the provisioning engine feature functionality
    /// </summary>
    [TestClass]
    public class FeatureNoScriptTests: FunctionalTestBase
    {
        #region Construction
        public FeatureNoScriptTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_28a75800-7295-4897-b8c8-4eb67cb2c553";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_28a75800-7295-4897-b8c8-4eb67cb2c553/sub";
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
#endif
}
