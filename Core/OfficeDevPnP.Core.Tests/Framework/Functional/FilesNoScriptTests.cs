using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !SP2013 && !SP2016
    [TestClass]
    public class FilesNoScriptTests : FunctionalTestBase
    {
        #region Construction
        public FilesNoScriptTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_da2a59c7-f789-4314-9889-2c57cb98d088";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_da2a59c7-f789-4314-9889-2c57cb98d088/sub";
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
        /// FilesTest Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionFilesTest()
        {
            new FilesImplementation().SiteCollectionFiles(centralSiteCollectionUrl);
        }

        /// <summary>
        /// Directory Files Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionDirectoryFilesTest()
        {
            new FilesImplementation().SiteCollectionDirectoryFiles(centralSiteCollectionUrl);
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
            new FilesImplementation().WebFiles(centralSiteCollectionUrl);
        }

        /// <summary>
        /// Directory Files Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebCollectionDirectoryFilesTest()
        {
            new FilesImplementation().WebDirectoryFiles(centralSiteCollectionUrl);
        }

        #endregion
    }
#endif
}
