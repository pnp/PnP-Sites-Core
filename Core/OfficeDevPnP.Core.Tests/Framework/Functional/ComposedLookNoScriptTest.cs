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
    public class ComposedLookNoScriptTest : FunctionalTestBase
    {
        #region Construction
        public ComposedLookNoScriptTest()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/bert1";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/bert1/sub";
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

        /// <summary>
        /// Site collection composed look test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionComposedLookTest()
        {
            new ComposedLookImplementation().SiteCollectionComposedLook(centralSiteCollectionUrl);
        }

        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebComposedLookTest()
        {
            new ComposedLookImplementation().WebComposedLook(centralSiteCollectionUrl, centralSubSiteUrl);
        }
    }
#endif
}
