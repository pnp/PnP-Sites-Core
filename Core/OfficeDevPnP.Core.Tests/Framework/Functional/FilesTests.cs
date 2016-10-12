using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
    public class FilesTests : FunctionalTestBase
    {
        #region Construction
        public FilesTests()
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
        /// FilesTest Test
        /// </summary>
        [TestMethod]
        public void FilesTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.HandlersToProcess = Handlers.Files;

                var result = TestProvisioningTemplate(cc, "files_add.xml", Handlers.Files, null, ptci);
                FilesValidator fv = new FilesValidator();
                Assert.IsTrue(fv.Validate(result.SourceTemplate.Files,cc));


            }
        }
        #endregion
    }
}
