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
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c89c25d3-4153-4464-8ad3-d0d6715fb6a8";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c89c25d3-4153-4464-8ad3-d0d6715fb6a8/sub";
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
        /// FilesTest Test
        /// </summary>
        [TestMethod]
        public void SiteCollectionFilesTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var result = TestProvisioningTemplate(cc, "files_add.xml", Handlers.Files);
                FilesValidator fv = new FilesValidator();
                Assert.IsTrue(fv.Validate(result.SourceTemplate.Files,cc));
            }
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// FilesTest Test
        /// </summary>
        [TestMethod]
        public void WebFilesTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                var result = TestProvisioningTemplate(cc, "files_add.xml", Handlers.Files);
                FilesValidator fv = new FilesValidator();
                Assert.IsTrue(fv.Validate(result.SourceTemplate.Files, cc));
            }
        }
        #endregion
    }
}
