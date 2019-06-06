using Microsoft.SharePoint.Client;
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

    public class SerializedContentType
    {
        public string SchemaXml { get; set; }
    }

    public class ContentTypeValidator : ValidatorBase
    {
        private bool isNoScriptSite = false;

        #region construction
        public ContentTypeValidator(Web web): base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:ContentTypes/pnp:ContentType";

            // Check if this is not a noscript site as we're not allowed to update some properties
            isNoScriptSite = web.IsNoScriptSite();
        }
        #endregion

        #region Validation logic
        public bool Validate(Core.Framework.Provisioning.Model.ContentTypeCollection sourceCollection, Core.Framework.Provisioning.Model.ContentTypeCollection targetCollection, TokenParser tokenParser)
        {
            // Convert object collections to XML 
            List<SerializedContentType> sourceContentTypes = new List<SerializedContentType>();
            List<SerializedContentType> targetContentTypes = new List<SerializedContentType>();

            foreach (Core.Framework.Provisioning.Model.ContentType ct in sourceCollection)
            {
                ProvisioningTemplate pt = new ProvisioningTemplate();
                pt.ContentTypes.Add(ct);

                sourceContentTypes.Add(new SerializedContentType() { SchemaXml = ExtractElementXml(pt) });                
            }

            foreach (Core.Framework.Provisioning.Model.ContentType ct in targetCollection)
            {
                ProvisioningTemplate pt = new ProvisioningTemplate();
                pt.ContentTypes.Add(ct);

                targetContentTypes.Add(new SerializedContentType() { SchemaXml = ExtractElementXml(pt) });
            }

            // Use XML validation logic to compare source and target
            Dictionary<string, string[]> parserSettings = new Dictionary<string, string[]>();
            parserSettings.Add("SchemaXml", null);
            bool isContentTypeMatch = ValidateObjectsXML(sourceContentTypes, targetContentTypes, "SchemaXml", new List<string> { "ID" }, tokenParser, parserSettings);
            Console.WriteLine("-- Content type validation " + isContentTypeMatch);
            return isContentTypeMatch;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            // Drop form URL properties if they're not provided in the source XML
            if (sourceObject.Attribute("NewFormUrl") == null)
            {
                DropAttribute(targetObject, "NewFormUrl");
            }
            if (sourceObject.Attribute("EditFormUrl") == null)
            {
                DropAttribute(targetObject, "EditFormUrl");
            }
            if (sourceObject.Attribute("DisplayFormUrl") == null)
            {
                DropAttribute(targetObject, "DisplayFormUrl");
            }
            if (sourceObject.Attribute("Description") == null)
            {
                DropAttribute(targetObject, "Description");
            }
            if (sourceObject.Attribute("Group") == null)
            {
                DropAttribute(targetObject, "Group");
            }

            // Since we can't upload aspx files there's no point in using the engine to set a custom content type forms
            if (isNoScriptSite)
            {
                DropAttribute(sourceObject, "NewFormUrl");
                DropAttribute(targetObject, "NewFormUrl");
                DropAttribute(sourceObject, "EditFormUrl");
                DropAttribute(targetObject, "EditFormUrl");
                DropAttribute(sourceObject, "DisplayFormUrl");
                DropAttribute(targetObject, "DisplayFormUrl");
            }

            // Target content type is retrieved with all fieldrefs, so delete the OOB ones
            XNamespace ns = SchemaVersion;
            var fieldRefs = targetObject.Descendants(ns + "FieldRefs").FirstOrDefault();

            IEnumerable<XElement> fieldRefElements = fieldRefs.Descendants(ns + "FieldRef");

            List<XElement> toDelete = new List<XElement>();

            foreach (var fieldRef in fieldRefElements)
            {
                // Delete the OOB fieldrefs
                if (BuiltInFieldId.Contains(new Guid(fieldRef.Attribute("ID").Value)) || 
                    fieldRef.Attribute("Name").Value.StartsWith("_dlc_") ||
                    fieldRef.Attribute("ID").Value.Equals("cbb92da4-fd46-4c7d-af6c-3128c2a5576e", StringComparison.InvariantCultureIgnoreCase) //DocumentSetDescription 
                    )
                {
                    toDelete.Add(fieldRef);
                }
                else
                {
                    // Drop the name attribute before the comparison
                    if (fieldRef.Attribute("Name") != null)
                    {
                        fieldRef.Attribute("Name").Remove();
                    }
                }
            }

            // Drop the OOB fieldrefs
            foreach (var fieldRef in toDelete)
            {
                fieldRef.Remove();
            }

            // Drop empty FieldRefs element
            fieldRefElements = fieldRefs.Descendants(ns + "FieldRef");
            if (!fieldRefElements.Any())
            {
                fieldRefs.Remove();
            }

            // Drop the WelcomePage attribute of the homepage
            if (sourceObject.Element(ns + "DocumentSetTemplate") != null)
            {
                if (sourceObject.Element(ns + "DocumentSetTemplate").Attribute("WelcomePage") != null)
                {
                    sourceObject.Element(ns + "DocumentSetTemplate").Attribute("WelcomePage").Remove();
                }
                
                if (isNoScriptSite)
                {
                    // Setting default documents is not supported in NoScript sites so let's drop that from the comparison
                    var defaultDocuments = sourceObject.Descendants(ns + "DefaultDocuments").FirstOrDefault();
                    if (defaultDocuments != null)
                    {
                        defaultDocuments.Remove();
                    }

                    defaultDocuments = targetObject.Descendants(ns + "DefaultDocuments").FirstOrDefault();
                    if (defaultDocuments != null)
                    {
                        defaultDocuments.Remove();
                    }
                }
                else
                {
                    // Drop the FileSourcePath attribute in both source and target
                    var defaultDocuments = targetObject.Descendants(ns + "DefaultDocuments").FirstOrDefault();
                    if (defaultDocuments != null)
                    {
                        IEnumerable<XElement> defaultDocumentsElements = defaultDocuments.Descendants(ns + "DefaultDocument");

                        foreach (var defaultDocument in defaultDocumentsElements)
                        {
                            DropAttribute(defaultDocument, "FileSourcePath");
                        }
                    }

                    defaultDocuments = sourceObject.Descendants(ns + "DefaultDocuments").FirstOrDefault();
                    if (defaultDocuments != null)
                    {
                        IEnumerable<XElement> defaultDocumentsElements = defaultDocuments.Descendants(ns + "DefaultDocument");

                        foreach (var defaultDocument in defaultDocumentsElements)
                        {
                            DropAttribute(defaultDocument, "FileSourcePath");
                        }
                    }
                }
            }

        }
        #endregion
    }
}
