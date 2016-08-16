using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional
{
    [TestClass]
    public class PropertyBagTests: FunctionalTestBase
    {

        #region Construction
        public PropertyBagTests()
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
        #endregion

        #region Site collection test cases
        [TestMethod]
        public void SiteCollectionPropertyBagAddingTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
            {
                // Ensure we can test clean
                DeleteWebProperties(cc);

                // Add web properties
                var result = TestProvisioningTemplate(cc, "propertybag_add.xml", Handlers.PropertyBagEntries);
                PropertyBagValidator pv = new PropertyBagValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.PropertyBagEntries, result.TargetTemplate.PropertyBagEntries, result.SourceTokenParser));

                // Update web properties
                var result2 = TestProvisioningTemplate(cc, "propertybag_delta_1.xml", Handlers.PropertyBagEntries);
                PropertyBagValidator pv2 = new PropertyBagValidator();
                pv2.ValidateEvent += Pv2_ValidateEvent;
                Assert.IsTrue(pv2.Validate(result2.SourceTemplate.PropertyBagEntries, result2.TargetTemplate.PropertyBagEntries, result2.SourceTokenParser));

                // Update system properties: run 1 is without specifying the override flag...no updates should happen
                ProvisioningTemplateApplyingInformation ptai3 = new ProvisioningTemplateApplyingInformation();
                ptai3.OverwriteSystemPropertyBagValues = false; //=default
                ptai3.HandlersToProcess = Handlers.PropertyBagEntries;
                // Set base template to null to ensure all properties are fetched by the engine
                ProvisioningTemplateCreationInformation ptci3 = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci3.BaseTemplate = null;
                ptci3.HandlersToProcess = Handlers.PropertyBagEntries;
                var result3 = TestProvisioningTemplate(cc, "propertybag_delta_2.xml", Handlers.PropertyBagEntries, ptai3, ptci3);
                PropertyBagValidator pv3 = new PropertyBagValidator();
                pv3.ValidateEvent += Pv3_ValidateEvent;
                Assert.IsTrue(pv3.Validate(result3.SourceTemplate.PropertyBagEntries, result3.TargetTemplate.PropertyBagEntries, result3.SourceTokenParser));

                // Update system properties: run 2 is with specifying the override flag...updates should happen if the overwrite flag was set to true
                ProvisioningTemplateApplyingInformation ptai4 = new ProvisioningTemplateApplyingInformation();
                // Set system overwrite flag
                ptai4.OverwriteSystemPropertyBagValues = true; 
                ptai4.HandlersToProcess = Handlers.PropertyBagEntries;
                // Set base template to null to ensure all properties are fetched by the engine
                ProvisioningTemplateCreationInformation ptci4 = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci4.BaseTemplate = null;
                ptci4.HandlersToProcess = Handlers.PropertyBagEntries;
                var result4 = TestProvisioningTemplate(cc, "propertybag_delta_2.xml", Handlers.PropertyBagEntries, ptai4, ptci4);
                PropertyBagValidator pv4 = new PropertyBagValidator();
                pv4.ValidateEvent += Pv2_ValidateEvent;
                Assert.IsTrue(pv4.Validate(result4.SourceTemplate.PropertyBagEntries, result4.TargetTemplate.PropertyBagEntries, result4.SourceTokenParser));
            }
        }
        #endregion

        #region Web test cases
        [TestMethod]
        public void WebPropertyBagAddingTest()
        {
            using (var cc = TestCommon.CreateClientContext(centralSubSiteUrl))
            {
                // Ensure we can test clean
                DeleteWebProperties(cc);

                // Add web properties
                var result = TestProvisioningTemplate(cc, "propertybag_add.xml", Handlers.PropertyBagEntries);
                PropertyBagValidator pv = new PropertyBagValidator();
                Assert.IsTrue(pv.Validate(result.SourceTemplate.PropertyBagEntries, result.TargetTemplate.PropertyBagEntries, result.SourceTokenParser));

                // Update web properties
                var result2 = TestProvisioningTemplate(cc, "propertybag_delta_1.xml", Handlers.PropertyBagEntries);
                PropertyBagValidator pv2 = new PropertyBagValidator();
                pv2.ValidateEvent += Pv2_ValidateEvent;
                Assert.IsTrue(pv2.Validate(result2.SourceTemplate.PropertyBagEntries, result2.TargetTemplate.PropertyBagEntries, result2.SourceTokenParser));

                // Update system properties: run 1 is without specifying the override flag...no updates should happen
                ProvisioningTemplateApplyingInformation ptai3 = new ProvisioningTemplateApplyingInformation();
                ptai3.OverwriteSystemPropertyBagValues = false; //=default
                ptai3.HandlersToProcess = Handlers.PropertyBagEntries;
                // Set base template to null to ensure all properties are fetched by the engine
                ProvisioningTemplateCreationInformation ptci3 = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci3.BaseTemplate = null;
                ptci3.HandlersToProcess = Handlers.PropertyBagEntries;
                var result3 = TestProvisioningTemplate(cc, "propertybag_delta_2.xml", Handlers.PropertyBagEntries, ptai3, ptci3);
                PropertyBagValidator pv3 = new PropertyBagValidator();
                pv3.ValidateEvent += Pv3_ValidateEvent;
                Assert.IsTrue(pv3.Validate(result3.SourceTemplate.PropertyBagEntries, result3.TargetTemplate.PropertyBagEntries, result3.SourceTokenParser));

                // Update system properties: run 2 is with specifying the override flag...updates should happen if the overwrite flag was set to true
                ProvisioningTemplateApplyingInformation ptai4 = new ProvisioningTemplateApplyingInformation();
                // Set system overwrite flag
                ptai4.OverwriteSystemPropertyBagValues = true;
                ptai4.HandlersToProcess = Handlers.PropertyBagEntries;
                // Set base template to null to ensure all properties are fetched by the engine
                ProvisioningTemplateCreationInformation ptci4 = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci4.BaseTemplate = null;
                ptci4.HandlersToProcess = Handlers.PropertyBagEntries;
                var result4 = TestProvisioningTemplate(cc, "propertybag_delta_2.xml", Handlers.PropertyBagEntries, ptai4, ptci4);
                PropertyBagValidator pv4 = new PropertyBagValidator();
                pv4.ValidateEvent += Pv2_ValidateEvent;
                Assert.IsTrue(pv4.Validate(result4.SourceTemplate.PropertyBagEntries, result4.TargetTemplate.PropertyBagEntries, result4.SourceTokenParser));
            }
        }
        #endregion

        #region Validation event handlers
        private void Pv2_ValidateEvent(object sender, ValidateEventArgs e)
        {
            // If "Overwrite" was set to false then we're not updating the property, hence we need to make an exception in our comparison logic
            if (e.PropertyName.Equals("Value", StringComparison.InvariantCultureIgnoreCase))
            {
                if ((e.SourceObject as PropertyBagEntry).Overwrite == false)
                {
                    // if source and target value are the same then somehow this delta update overwrote which it shouldn't
                    if (!e.SourceValue.Equals(e.TargetValue))
                    {
                        e.IsEqual = true;
                    }
                }
            }
        }

        private void Pv3_ValidateEvent(object sender, ValidateEventArgs e)
        {
            // We didn't specify the flag that allows system properties to be updated, all value's should be different
            if (e.PropertyName.Equals("Value", StringComparison.InvariantCultureIgnoreCase))
            {
                if (!e.SourceValue.Equals(e.TargetValue))
                {
                    e.IsEqual = true;
                }
            }
        }
        #endregion

        #region Helper methods
        private void DeleteWebProperties(ClientContext cc)
        {
            cc.Web.AllProperties.ClearObjectData();

            var props = cc.Web.AllProperties;
            cc.Web.Context.Load(props);
            cc.Web.Context.ExecuteQueryRetry();

            List<string> propsToRemove = new List<string>();
            foreach(var prop in props.FieldValues)
            {
                if (prop.Key.StartsWith("PROP_"))
                {
                    propsToRemove.Add(prop.Key);
                }
            }

            foreach(var prop in propsToRemove)
            {
                cc.Web.RemovePropertyBagValue(prop);
            }
        }
        #endregion
    }
}
