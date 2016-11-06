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
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/bert1";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/bert1/sub";
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
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionComposedLookTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                if (!cc.Web.IsNoScriptSite())
                {
                    // Add supporting files
                    TestProvisioningTemplate(cc, "composedlook_files.xml", Handlers.Files);

                    var result = TestProvisioningTemplate(cc, "composedlook_add_1.xml", Handlers.ComposedLook);
                    ComposedLookValidator composedLookVal = new ComposedLookValidator();
                    Assert.IsTrue(composedLookVal.Validate(result.SourceTemplate.ComposedLook, result.TargetTemplate.ComposedLook));

                    var result2 = TestProvisioningTemplate(cc, "composedlook_add_2.xml", Handlers.ComposedLook);
                    Assert.IsTrue(composedLookVal.Validate(result2.SourceTemplate.ComposedLook, result2.TargetTemplate.ComposedLook));
                }
            }
        }

        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebComposedLookTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                // Add supporting files
                TestProvisioningTemplate(cc, "composedlook_files.xml", Handlers.Files);
            }

            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                if (!cc.Web.IsNoScriptSite())
                {
                    var result = TestProvisioningTemplate(cc, "composedlook_add_1.xml", Handlers.ComposedLook);
                    ComposedLookValidator composedLookVal = new ComposedLookValidator();
                    Assert.IsTrue(composedLookVal.Validate(result.SourceTemplate.ComposedLook, result.TargetTemplate.ComposedLook));

                    var result2 = TestProvisioningTemplate(cc, "composedlook_add_2.xml", Handlers.ComposedLook);
                    Assert.IsTrue(composedLookVal.Validate(result2.SourceTemplate.ComposedLook, result2.TargetTemplate.ComposedLook));
                }
            }
        }

    }
}
