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
    [TestClass]
    public class WorkflowsTests : FunctionalTestBase
    {
        #region Construction
        public WorkflowsTests()
        {
            //debugMode = true;
            centralSiteCollectionUrl = "https://crtlab2.sharepoint.com/sites/source2";
            centralSubSiteUrl = "https://crtlab2.sharepoint.com/sites/source2/sub2";
        }
        #endregion

        #region Test setup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            //ClassInitBase(context);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            //ClassCleanupBase();
        }
        #endregion

        #region Site collection test cases
        /// <summary>
        /// WorkflowsTests Test
        /// </summary>
        [TestMethod]
        public void WorkflowsTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.HandlersToProcess = Handlers.Workflows;
                ptci.FileConnector= new FileSystemConnector(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");


                var result = TestProvisioningTemplate(cc, "workflows_add.xml", Handlers.Workflows,null,ptci);
                WorkflowValidator wv = new WorkflowValidator();
                Assert.IsTrue(wv.Validate(result.SourceTemplate.Workflows,result.TargetTemplate.Workflows,result.TargetTokenParser));
            }
        }
        #endregion
    }
}
