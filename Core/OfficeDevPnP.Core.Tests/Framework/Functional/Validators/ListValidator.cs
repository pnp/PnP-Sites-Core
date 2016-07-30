using Newtonsoft.Json;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{

    public class SerializedListInstance
    {
        public string SchemaXml { get; set; }
    }

    public class ListInstanceValidator : ValidatorBase
    {
        #region construction
        public ListInstanceValidator(): base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:Lists/pnp:ListInstance";
        }
        #endregion

        #region Validation logic
        public bool Validate(ListInstanceCollection sourceCollection, ListInstanceCollection targetCollection, TokenParser tokenParser)
        {
            // Convert object collections to XML 
            List<SerializedListInstance> sourceLists = new List<SerializedListInstance>();
            List<SerializedListInstance> targetLists = new List<SerializedListInstance>();

            foreach (ListInstance list in sourceCollection)
            {
                // don't add hidden lists since they're not exported again...
                if (!list.Hidden)
                {
                    ProvisioningTemplate pt = new ProvisioningTemplate();
                    pt.Lists.Add(list);

                    sourceLists.Add(new SerializedListInstance() { SchemaXml = ExtractElementXml(pt) });
                }
            }

            foreach (ListInstance list in targetCollection)
            {
                ProvisioningTemplate pt = new ProvisioningTemplate();
                pt.Lists.Add(list);

                targetLists.Add(new SerializedListInstance() { SchemaXml = ExtractElementXml(pt) });
            }

            // Use XML validation logic to compare source and target
            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            parserSettings.Add("SchemaXml", null);
            bool isListsMatch = ValidateObjectsXML(sourceLists, targetLists, "SchemaXml", new List<string> { "Title" }, tokenParser, parserSettings);
            Console.WriteLine("-- Lists validation " + isListsMatch);
            return isListsMatch;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            // Drop list properties if they're not provided in the source XML
            string[] ListProperties = new string[] { "Description", "DocumentTemplate", "MinorVersionLimit", "MaxVersionLimit", "DraftVersionVisibility", "TemplateFeatureID", "EnableAttachments" };
            foreach (string listProperty in ListProperties)
            {
                if (sourceObject.Attribute(listProperty) == null)
                {
                    DropAttribute(targetObject, listProperty);
                }
            }

            // Drop list elements if they're not provided in the source XML
            string[] ListElements = new string[] { "ContentTypeBindings", "Views", "FieldRefs", "Fields" };
            foreach (var listElement in ListElements)
            {
                var sourceListElementXML = sourceObject.Descendants(ns + listElement);
                var targetListElementXML = targetObject.Descendants(ns + listElement);
                if (sourceListElementXML.Any() == false && targetListElementXML.Any() == true)
                {
                    targetListElementXML.Remove();
                }
            }

#if ONPREMISES
            // MaxVersionLimit cannot be set in on-premises, so remove it before comparing before and after
            if (sourceObject.Attribute("MaxVersionLimit") != null)
            {
                DropAttribute(targetObject, "MaxVersionLimit");
                DropAttribute(sourceObject, "MaxVersionLimit");
            }
#endif

            // If RemoveExistingContentTypes is set then remove the attribute from source since on target we don't add this. 
            var contentTypeBindings = targetObject.Descendants(ns + "ContentTypeBinding");
            bool removeExistingContentTypes = false;
            if (sourceObject.Attribute("RemoveExistingContentTypes") != null)
            {
                removeExistingContentTypes = sourceObject.Attribute("RemoveExistingContentTypes").Value.ToBoolean();
                DropAttribute(sourceObject, "RemoveExistingContentTypes");
            }

            if (contentTypeBindings != null && contentTypeBindings.Any())
            {
                // One can add ContentTypeBindings without specifying ContentTypesEnabled. The engine will turn on ContentTypesEnabled automatically in that case
                if (sourceObject.Attribute("ContentTypesEnabled") == null)
                {
                    DropAttribute(targetObject, "ContentTypesEnabled");
                }

                if (removeExistingContentTypes)
                {
                    foreach (var contentTypeBinding in contentTypeBindings.ToList())
                    {
                        // Remove the folder content type if present because we're not removing that one via RemoveExistingContentTypes
                        if (contentTypeBinding.Attribute("ContentTypeID").Value == "0x0120")
                        {
                            contentTypeBinding.Remove();
                        }
                    }
                }
                else // We did not remove existing content types
                {
                    var sourceContentTypeBindings = sourceObject.Descendants(ns + "ContentTypeBinding");
                    foreach (var contentTypeBinding in contentTypeBindings.ToList())
                    {
                        string value = contentTypeBinding.Attribute("ContentTypeID").Value;
                        // drop all content types which are not mentioned in the source
                        if (sourceContentTypeBindings.Where(p => p.Attribute("ContentTypeID").Value == value).Any() == false)
                        {
                            contentTypeBinding.Remove();
                        }                   
                    }
                }
            }
        }
        #endregion
    }
}
