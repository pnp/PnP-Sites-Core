using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class CustomActionTests: FunctionalTestBase
    {

        #region Construction
        public CustomActionTests()
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
        public void SiteCollectionCustomActionAddingTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                // Ensure we can test clean
                DeleteCustomActions(cc);

                // Add custom actions
                var result = TestProvisioningTemplate(cc, "customaction_add.xml", Handlers.CustomActions);
                Assert.IsTrue(CustomActionValidator.Validate(result.SourceTemplate.CustomActions, result.TargetTemplate.CustomActions, result.TargetTokenParser));

#if !SP2013
                // Update custom actions
                var result2 = TestProvisioningTemplate(cc, "customaction_delta_1.xml", Handlers.CustomActions);
                Assert.IsTrue(CustomActionValidator.Validate(result2.SourceTemplate.CustomActions, result2.TargetTemplate.CustomActions, result2.TargetTokenParser));

                // Update custom actions
                var result3 = TestProvisioningTemplate(cc, "customaction_1605_delta_2.xml", Handlers.CustomActions);
                Assert.IsTrue(CustomActionValidator.Validate(result3.SourceTemplate.CustomActions, result3.TargetTemplate.CustomActions, result3.TargetTokenParser));
#endif
            }
        }
        #endregion

        #region Web test cases
        [TestMethod]
        public void WebCustomActionAddingTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                // Ensure we can test clean
                DeleteCustomActions(cc);

                // Add custom actions
                var result = TestProvisioningTemplate(cc, "customaction_add.xml", Handlers.CustomActions);
                Assert.IsTrue(CustomActionValidator.ValidateCustomActions(result.SourceTemplate.CustomActions.WebCustomActions, result.TargetTemplate.CustomActions.WebCustomActions, result.TargetTokenParser));

#if !SP2013
                // Update custom actions
                var result2 = TestProvisioningTemplate(cc, "customaction_delta_1.xml", Handlers.CustomActions);
                Assert.IsTrue(CustomActionValidator.ValidateCustomActions(result2.SourceTemplate.CustomActions.WebCustomActions, result2.TargetTemplate.CustomActions.WebCustomActions, result2.TargetTokenParser));
#endif
            }
        }
#endregion

#region Helper methods
        private void DeleteCustomActions(ClientContext cc)
        {
            var siteActions = cc.Site.GetCustomActions();
            foreach (var action in siteActions)
            {
                if (action.Name.StartsWith("CA_"))
                {
                    cc.Site.DeleteCustomAction(action.Id);
                }
            }

            var webActions = cc.Web.GetCustomActions();
            foreach (var action in webActions)
            {
                if (action.Name.StartsWith("CA_"))
                {
                    cc.Web.DeleteCustomAction(action.Id);
                }
            }
        }
#endregion
    }
}
