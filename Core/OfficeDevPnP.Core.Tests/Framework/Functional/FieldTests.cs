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
            debugMode = true;
            centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59";
            centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c3a9328a-21dd-4d3e-8919-ee73b0d5db59/sub";
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
        [TestMethod]
        public void SiteCollectionFieldAddingTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                // Ensure we can test clean
                DeleteFields(cc);

                // Add web properties
                var result = TestProvisioningTemplate(cc, "field_add.xml", Handlers.Fields | Handlers.TermGroups);
                FieldValidator pv = new FieldValidator();
                pv.ValidateXmlEvent += Pv_ValidateXmlEvent;              
                Assert.IsTrue(pv.Validate(result.SourceTemplate.SiteFields, result.TargetTemplate.SiteFields, result.TargetTokenParser));
            }
        }

        private void Pv_ValidateXmlEvent(object sender, ValidateXmlEventArgs e)
        {
            string fieldType = e.SourceObject.Attribute("Type").Value;

            // Drop attributes before comparison
            DropAttribute(e.TargetObject, "SourceID");
            // If not StaticName was presented in the source it's returned in the target anyhow, so let's drop it before comparison
            if (e.SourceObject.Attribute("StaticName") == null)
            {
                DropAttribute(e.TargetObject, "StaticName");
            }

            // Fix ID attribute value casing before comparison
            UpperCaseAttribute(e.TargetObject, "ID");
            UpperCaseAttribute(e.SourceObject, "ID");

            // Drop the target FieldRefs element for calculated fields as that's the way how the engine extracts these fields
            if (fieldType.Equals("Calculated", StringComparison.InvariantCultureIgnoreCase))
            {
                var formulaNode = e.SourceObject.Descendants("Formula").FirstOrDefault();
                if (formulaNode != null)
                {
                    var fieldRefs = e.SourceObject.Descendants("FieldRefs");
                    if (fieldRefs != null)
                    {
                        fieldRefs.Remove();
                    }
                }
            }
            else if (fieldType.Equals("TaxonomyFieldType", StringComparison.InvariantCultureIgnoreCase))
            {
                // TaxonomyFieldType has specific validation behaviour

                // Step 1: Drop List and WebId attributes
                DropAttribute(e.TargetObject, "List");
                DropAttribute(e.TargetObject, "WebId");

                // Step 2: Compare the customization properties
                string[] customizationProperties = new string[] { "GroupId", "AnchorId", "UserCreated", "Open", "TextField", "IsPathRendered", "IsKeyword", "CreateValuesInEditForm", "FilterAssemblyStrongName", "FilterClassName", "FilterMethodName", "FilterJavascriptProperty", "TargetTemplate" };
                bool customizationPropertiesAreEqual = true;
                foreach (string customizationProperty in customizationProperties)
                {
                    if (!TaxonomyFieldCustomizationPropertyIsEqual(e.SourceObject, e.TargetObject, customizationProperty))
                    {
                        customizationPropertiesAreEqual = false;
                        break;
                    }
                }

                // Step 3: if customization properties are equal then drop them so that the xml comparison can be done
                if (customizationPropertiesAreEqual)
                {
                    // drop the customization element so that the base implementation using the XNode.DeepComparison works
                    var customizationXml = e.TargetObject.Descendants("Customization");
                    if (customizationXml != null)
                    {
                        customizationXml.Remove();
                    }

                    customizationXml = e.SourceObject.Descendants("Customization");
                    if (customizationXml != null)
                    {
                        customizationXml.Remove();
                    }
                }
                else
                {
                    // let the deep comparison fail...
                }
            }
        }


        #endregion

        #region Web test cases

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
