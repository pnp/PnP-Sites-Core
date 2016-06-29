using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class FeatureTests: FunctionalTestBase
    {
        public FeatureTests()
        {
            debugMode = true;
            centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59";
            centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59/sub";
        }

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

        [TestMethod]
        public void SiteCollectionTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var result = ApplyProvisioningTemplate(cc, "feature_base.xml");

                Assert.IsTrue(FeatureValidator.Validate(result.Item1.Features, result.Item2.Features));
            }
        }

        [TestMethod]
        public void WebTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                var result = ApplyProvisioningTemplate(cc, "feature_base.xml");
                Assert.IsTrue(FeatureValidator.ValidateFeatures(result.Item1.Features.WebFeatures, result.Item2.Features.WebFeatures));
            }
        }

        //cc.Load(cc.Web, w => w.Title);
        //cc.ExecuteQuery();

        //using (var ccSub = cc.Clone(centralSubSiteUrl))
        //{
        //    ccSub.Load(ccSub.Web, w => w.Title);
        //    ccSub.ExecuteQuery();
        //}


    }
}
