using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

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
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59/sub";
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
        public void SiteCollectionFeatureActivationDeactivationTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var result = TestProvisioningTemplate(cc, "feature_base.xml", Handlers.Features);
                Assert.IsTrue(FeatureValidator.Validate(result.SourceTemplate.Features, result.TargetTemplate.Features));
            }
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Verifies is the web scoped features are correctly activated/deactivated for a sub site
        /// </summary>
        [TestMethod]
        public void WebFeatureActivationDeactivationTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                var result = TestProvisioningTemplate(cc, "feature_base.xml", Handlers.Features);
                Assert.IsTrue(FeatureValidator.ValidateFeatures(result.SourceTemplate.Features.WebFeatures, result.TargetTemplate.Features.WebFeatures));
            }
        }
        #endregion
    }
}
