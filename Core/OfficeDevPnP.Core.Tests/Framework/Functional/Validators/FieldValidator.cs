using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    class FieldValidator : ValidatorBase
    {
        #region construction
        public FieldValidator() : base()
        {
            // optionally override schema version
            // SchemaVersion = "http://schemas.dev.office.com/PnP/2016/05/ProvisioningSchema";
            // XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:ContentTypes/pnp:ContentType";
        }
        #endregion

        #region Validation logic
        public bool Validate(FieldCollection sourceCollection, FieldCollection targetCollection, TokenParser tokenParser)
        {
            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            parserSettings.Add("SchemaXml", new string[] { "~sitecollection", "~site", "{sitecollectiontermstoreid}", "{termsetid}" });
            bool isFieldMatch = ValidateObjectsXML(sourceCollection, targetCollection, "SchemaXml", new List<string> { "ID" }, tokenParser, parserSettings);
            Console.WriteLine("-- Field validation " + isFieldMatch);
            return isFieldMatch;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            string fieldType = sourceObject.Attribute("Type").Value;

            // Drop attributes before comparison
            DropAttribute(targetObject, "SourceID");
            // If not StaticName was presented in the source it's returned in the target anyhow, so let's drop it before comparison
            if (sourceObject.Attribute("StaticName") == null)
            {
                DropAttribute(targetObject, "StaticName");
            }

            // Fix ID attribute value casing before comparison
            UpperCaseAttribute(targetObject, "ID");
            UpperCaseAttribute(sourceObject, "ID");

            if (fieldType.Equals("Calculated", StringComparison.InvariantCultureIgnoreCase))
            {
                // Calculated has specific validation behaviour

                var formulaNode = sourceObject.Descendants("Formula").FirstOrDefault();
                if (formulaNode != null)
                {
                    // The engine drops the FieldRefs element when providing a calculated field
                    var fieldRefs = sourceObject.Descendants("FieldRefs");
                    if (fieldRefs != null)
                    {
                        fieldRefs.Remove();
                    }

                    // Dropping Formula elements since the engine is creating (tokenized) formula's that use the display name instead of the Name property
                    formulaNode.Remove();
                    formulaNode = targetObject.Descendants("Formula").FirstOrDefault();
                    if (formulaNode != null)
                    {
                        formulaNode.Remove();
                    }
                }
            }
            else if (fieldType.Equals("TaxonomyFieldType", StringComparison.InvariantCultureIgnoreCase) || fieldType.Equals("TaxonomyFieldTypeMulti", StringComparison.InvariantCultureIgnoreCase))
            {
                // TaxonomyFieldType has specific validation behaviour

                // Step 1: Drop List and WebId attributes
                DropAttribute(targetObject, "List");
                DropAttribute(targetObject, "WebId");

                // Step 2: Compare the customization properties
                string[] customizationProperties = new string[] { "GroupId", "AnchorId", "UserCreated", "Open", "TextField", "IsPathRendered", "IsKeyword", "CreateValuesInEditForm", "FilterAssemblyStrongName", "FilterClassName", "FilterMethodName", "FilterJavascriptProperty", "TargetTemplate" };
                bool customizationPropertiesAreEqual = true;
                foreach (string customizationProperty in customizationProperties)
                {
                    if (!TaxonomyFieldCustomizationPropertyIsEqual(sourceObject, targetObject, customizationProperty))
                    {
                        customizationPropertiesAreEqual = false;
                        break;
                    }
                }

                // Step 3: if customization properties are equal then drop them so that the xml comparison can be done
                if (customizationPropertiesAreEqual)
                {
                    // drop the customization elements so that the base XML comparison implementation works
                    var customizationXml = targetObject.Descendants("Customization");
                    if (customizationXml != null)
                    {
                        customizationXml.Remove();
                    }

                    customizationXml = sourceObject.Descendants("Customization");
                    if (customizationXml != null)
                    {
                        customizationXml.Remove();
                    }
                }
                else
                {
                    // let the xml comparison fail...
                }
            }
        }

        #endregion
    }
}
