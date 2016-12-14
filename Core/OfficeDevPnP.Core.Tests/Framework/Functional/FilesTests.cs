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
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c81e4b0d-0242-4c80-8272-18f13e759333";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c81e4b0d-0242-4c80-8272-18f13e759333/sub";
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
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionFilesTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                var result = TestProvisioningTemplate(cc, "files_add.xml", Handlers.Files | Handlers.Lists);
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
        [Timeout(15 * 60 * 1000)]
        public void WebFilesTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                // Ensure we can test clean
                DeleteLists(cc);

                var result = TestProvisioningTemplate(cc, "files_add.xml", Handlers.Files | Handlers.Lists);
                FilesValidator fv = new FilesValidator();
                Assert.IsTrue(fv.Validate(result.SourceTemplate.Files, cc));
            }
        }
        #endregion

        #region Helper methods
        private void DeleteLists(ClientContext cc)
        {
            DeleteListsImplementation(cc);
        }

        private static void DeleteListsImplementation(ClientContext cc)
        {
            cc.Load(cc.Web.Lists, f => f.Include(t => t.Title));
            cc.ExecuteQueryRetry();

            foreach (var list in cc.Web.Lists.ToList())
            {
                if (list.Title.StartsWith("LI_"))
                {
                    list.DeleteObject();
                }
            }
            cc.ExecuteQueryRetry();
        }
        #endregion
    }
}
