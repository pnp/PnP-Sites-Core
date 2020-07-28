using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Implementation;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Test cases for the provisioning engine Publishing functionality
    /// </summary>
    [TestClass]
    public class PublishingNoScriptTest : FunctionalTestBase
    {
        #region Construction
        public PublishingNoScriptTest()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_a21016d5-886f-49eb-9984-f9f4dce76741";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_a21016d5-886f-49eb-9984-f9f4dce76741/sub";
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
        /// Design package publishing test in site collection
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionPublishingTest()
        {
            new PublishingImplementation().SiteCollectionPublishing(centralSiteCollectionUrl);
        }
        #endregion        
    }
#endif
}
