using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class CustomActionTests: FunctionalTestBase
    {

        #region Construction
        public CustomActionTests()
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

        #region debug
        [TestMethod]
        public void DebugTest()
        {
            bool cleanup = true;

            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var siteActions = cc.Site.GetCustomActions();
                foreach (var action in siteActions)
                {
                    if (action.Name.StartsWith("CA_"))
                    {
                        if (cleanup)
                        {
                            cc.Site.DeleteCustomAction(action.Id);
                        }
                        else
                        {
                            if (action.Name == "CA_1")
                            {
                                action.Description = "Site Classification Application - Updated";
                                action.Group = "SiteActions";
                                action.Location = "ScriptLink";
                                action.Title = "Site Classification for Bert";
                                action.Sequence = 1001;
                                action.ScriptBlock = "alert('hi Bert');";
                                action.Url = null;
                                action.Update();
                                cc.ExecuteQueryRetry();
                            }
                        }
                    }
                }

                var webActions = cc.Web.GetCustomActions();
                foreach (var action in webActions)
                {
                    if (action.Name.StartsWith("CA_"))
                    {
                        if (cleanup)
                        {
                            cc.Web.DeleteCustomAction(action.Id);
                        }
                        else
                        {

                        }
                    }
                }
            }

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

                // Update custom actions
                var result2 = TestProvisioningTemplate(cc, "customaction_delta_1.xml", Handlers.CustomActions);
                Assert.IsTrue(CustomActionValidator.Validate(result2.SourceTemplate.CustomActions, result2.TargetTemplate.CustomActions, result2.TargetTokenParser));
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

                // Update custom actions
                var result2 = TestProvisioningTemplate(cc, "customaction_delta_1.xml", Handlers.CustomActions);
                Assert.IsTrue(CustomActionValidator.ValidateCustomActions(result2.SourceTemplate.CustomActions.WebCustomActions, result2.TargetTemplate.CustomActions.WebCustomActions, result2.TargetTokenParser));
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
