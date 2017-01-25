using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605;
using ContentType = OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType;
using OfficeDevPnP.Core.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Serialization;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers;
using FileLevel = OfficeDevPnP.Core.Framework.Provisioning.Model.FileLevel;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    internal class XMLPnPSchemaV201605Serializer : IXMLSchemaFormatter, ITemplateFormatter
    {
        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        string IXMLSchemaFormatter.NamespaceUri
        {
            get { return (XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05); }
        }

        string IXMLSchemaFormatter.NamespacePrefix
        {
            get { return (XMLConstants.PROVISIONING_SCHEMA_PREFIX); }
        }

        public bool IsValid(Stream template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            // Load the template into an XDocument
            XDocument xml = XDocument.Load(template);

            // Load the XSD embedded resource
            Stream stream = typeof(XMLPnPSchemaV201605Serializer)
                .Assembly
                .GetManifestResourceStream("OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2016-05.xsd");

            // Prepare the XML Schema Set
            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add(XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05,
                new XmlTextReader(stream));

            Boolean result = true;
            xml.Validate(schemas, (o, e) =>
            {
                Diagnostics.Log.Error(e.Exception, "SchemaFormatter", "Template is not valid: {0}", e.Message);
                result = false;
            });

            return (result);
        }

        Stream ITemplateFormatter.ToFormattedTemplate(Model.ProvisioningTemplate template)
        {
            return null;
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(Stream template, String identifier)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            // Crate a copy of the source stream
            MemoryStream sourceStream = new MemoryStream();
            template.CopyTo(sourceStream);
            sourceStream.Position = 0;

            // Check the provided template against the XML schema
            if (!this.IsValid(sourceStream))
            {
                // TODO: Use resource file
                throw new ApplicationException("The provided template is not valid!");
            }

            sourceStream.Position = 0;
            XDocument xml = XDocument.Load(sourceStream);
            XNamespace pnp = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;


            // Prepare a variable to hold the single source formatted template
            //V201605.ProvisioningTemplate source = null;

            // Prepare a variable to hold the resulting ProvisioningTemplate instance
            Model.ProvisioningTemplate result = new Model.ProvisioningTemplate();

            // Determine if we're working on a wrapped SharePointProvisioningTemplate or not
            if (xml.Root.Name == pnp + "Provisioning")
            {

                var rootElement = xml.Root;

                // Deserialize preferences
                var preferences = rootElement.Elements(pnp + "Preferences");
                if (preferences != null)
                {
                    foreach (var preference in preferences)
                    {
                        var parameters = preferences.Elements(pnp + "Parameters");
                        foreach (var parameter in parameters)
                        {
                            var key = parameter.Attribute("Key").Value;
                            var value = parameter.Value;
                            result.Parameters.Add(key, value);
                        }
                    }
                }

                // TODO: Localizations

                var templatesElement = rootElement.Elements(pnp + "Templates").FirstOrDefault();

                if (templatesElement != null)
                {
                    var provisioningTemplatesElements = templatesElement.Elements(pnp + "ProvisioningTemplate");

                    foreach (var templateElement in provisioningTemplatesElements)
                    {
                        // TODO : handle the in-place templates
                        // TODO : handle external file templates
                    }

                    var provisioningTemplateElement = provisioningTemplatesElements.FirstOrDefault();
                    if (provisioningTemplateElement != null)
                    {

                        var idAttribute = provisioningTemplateElement.Attribute(pnp + "ID");
                        result.Id = idAttribute?.Value;

                        // Get all serializers
                        var currentAssembly = this.GetType().Assembly;
                        var serializers = currentAssembly.GetTypes()
                            .Where(t => t.GetInterface(typeof(ISchemaSerializer).FullName) != null)
                            .Where(t =>
                            {
                                var attribute = t.GetCustomAttributes<SupportedTemplateSchemasAttribute>(false).FirstOrDefault();
                                return attribute != null && attribute.Schemas.HasFlag(SupportedSchema.V201605);
                            }
                          );
                        foreach (var serializer in serializers)
                        {
                            bool execute = true;

                            if (execute)
                            {
                                var instance = Activator.CreateInstance(serializer) as ISchemaSerializer;
                                if (instance != null)
                                {
                                    result = instance.ToProvisioningTemplate(provisioningTemplateElement, pnp, result);
                                }
                            }
                        }

                    }
                }


            }
            else if (xml.Root.Name == pnp + "ProvisioningTemplate")
            {
                //var IdAttribute = xml.Root.Attribute("ID");

                //// If there is a provided ID, and if it doesn't equal the current ID
                //if (!String.IsNullOrEmpty(identifier) &&
                //    IdAttribute != null &&
                //    IdAttribute.Value != identifier)
                //{
                //    // TODO: Use resource file
                //    throw new ApplicationException("The provided template identifier is not available!");
                //}
                //else
                //{
                //    source = XMLSerializer.Deserialize<V201605.ProvisioningTemplate>(xml);
                //}
            }

            return (result);
        }
    }
}

