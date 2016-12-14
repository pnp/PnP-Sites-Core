using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class FieldTests : FunctionalTestBase
    {
        #region Construction
        public FieldTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59/sub";
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

            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Test that require taxonomy creation are not supported in app-only.");
            }
        }
        #endregion

        #region Site collection test cases
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionFieldAddingTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                // Ensure we can test clean
                DeleteFields(cc);

                // Add fields
                var result = TestProvisioningTemplate(cc, "field_add.xml", Handlers.Fields | Handlers.TermGroups);
                FieldValidator pv = new FieldValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.SiteFields, result.TargetTemplate.SiteFields, result.TargetTokenParser));

                // Apply delta to fields
                var result2 = TestProvisioningTemplate(cc, "field_delta_1.xml", Handlers.Fields);
                Assert.IsTrue(pv.Validate(result2.SourceTemplate.SiteFields, result2.TargetTemplate.SiteFields, result2.TargetTokenParser));
            }
        }
        #endregion

        #region Web test cases
        // No need to have these as the engine is blocking creation and extraction of fields at web level
        #endregion

        #region Validation event handlers
        #endregion

        #region Helper methods
        private void DeleteFields(ClientContext cc)
        {
            cc.Load(cc.Web.Fields, f => f.Include(t => t.InternalName));
            cc.ExecuteQueryRetry();

            foreach (var field in cc.Web.Fields.ToList())
            {
                // First drop the fields that have 2 _'s...convention used to name the fields dependent on a lookup.
                if (field.InternalName.Replace("FLD_", "").IndexOf("_") > 0)
                {
                    if (field.InternalName.StartsWith("FLD_"))
                    {
                        field.DeleteObject();
                    }
                }
            }

            foreach (var field in cc.Web.Fields.ToList())
            {
                if (field.InternalName.StartsWith("FLD_"))
                {
                    field.DeleteObject();
                }
            }

            cc.ExecuteQueryRetry();
            
        }

        #endregion
    }
}
