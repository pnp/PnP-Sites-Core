using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !ONPREMISES
    [TestClass]
    public class WorkflowsTests : FunctionalTestBase
    {
        #region Construction
        public WorkflowsTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_60453ae9-a218-436e-9231-cb9da3c4fdd3";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_60453ae9-a218-436e-9231-cb9da3c4fdd3/sub";
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

            if (new Uri(TestCommon.DevSiteUrl).DnsSafeHost.Contains("spoppe.com")) 
            {
                Assert.Inconclusive("Test that require workflow can't be running on edog.");
            }
        }
        #endregion

        #region Site collection test cases
        /// <summary>
        /// WorkflowsTests Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionWorkflowsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.HandlersToProcess = Handlers.Lists | Handlers.Workflows;
                ptci.FileConnector= new FileSystemConnector(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");


                var result = TestProvisioningTemplate(cc, "workflows_add_1605.xml", Handlers.Lists | Handlers.Workflows, null, ptci);
                WorkflowValidator wv = new WorkflowValidator();
                Assert.IsTrue(wv.Validate(result.SourceTemplate.Workflows, result.TargetTemplate.Workflows, result.TargetTokenParser));
            }
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// WorkflowsTests Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebWorkflowsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.HandlersToProcess = Handlers.Lists | Handlers.Workflows;
                ptci.FileConnector = new FileSystemConnector(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");


                var result = TestProvisioningTemplate(cc, "workflows_add_1605.xml", Handlers.Lists | Handlers.Workflows, null, ptci);
                WorkflowValidator wv = new WorkflowValidator();
                Assert.IsTrue(wv.Validate(result.SourceTemplate.Workflows, result.TargetTemplate.Workflows, result.TargetTokenParser));
            }
        }
        #endregion
    }
#endif
}
