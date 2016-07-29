using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine Publishing functionality
    /// </summary>
    [TestClass]
    public class ComposedLookTest : FunctionalTestBase
    {
        #region Construction
        public ComposedLookTest()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://crtlab2.sharepoint.com/sites/source2";
            //centralSubSiteUrl = "https://crtlab2.sharepoint.com/sites/source2/sub2";
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

        /// <summary>
        /// Site collection composed look test
        /// </summary>
        [TestMethod]
        public void SiteCollectionComposedLookTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                var result = TestProvisioningTemplate(cc, "composedlook_add.xml", Handlers.ComposedLook);
                ComposedLookValidator composedLookVal = new ComposedLookValidator();
                Assert.IsTrue(composedLookVal.Validate(result.SourceTemplate.ComposedLook, result.TargetTemplate.ComposedLook));
            }
        }

        [TestMethod]
        public void WebComposedLookTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                var result = TestProvisioningTemplate(cc, "composedlook_add.xml", Handlers.ComposedLook);
                ComposedLookValidator composedLookVal = new ComposedLookValidator();
                Assert.IsTrue(composedLookVal.Validate(result.SourceTemplate.ComposedLook, result.TargetTemplate.ComposedLook));
            }
        }

    }
}
