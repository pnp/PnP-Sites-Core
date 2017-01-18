using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;
using ContentType = OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType;
using FileLevel = OfficeDevPnP.Core.Framework.Provisioning.Model.FileLevel;
using AutoMapper;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperProfiles;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    internal abstract class XMLPnPSchemaAutoMapperFormatter : 
        ITemplateFormatter, IXMLSchemaFormatter
    {
        public XMLPnPSchemaAutoMapperFormatter(String xmlNamespace, Stream xmlSchema)
        {
            this._xmlNamespace = xmlNamespace;
            this._xmlSchema = xmlSchema;
        }

        private String _xmlNamespace;
        private Stream _xmlSchema;

        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        string IXMLSchemaFormatter.NamespaceUri
        {
            get { return (this._xmlNamespace); }
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

            // Prepare the XML Schema Set
            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add(this._xmlNamespace, new XmlTextReader(this._xmlSchema));

            Boolean result = true;
            xml.Validate(schemas, (o, e) =>
            {
                Diagnostics.Log.Error(e.Exception, "SchemaFormatter", "Template is not valid: {0}", e.Message);
                result = false;
            });

            return (result);
        }

        /// <summary>
        /// Abstract method used to prepare a Mapper object ready to map from Domain Model to the XML object model
        /// </summary>
        /// <returns></returns>
        protected abstract IMapper CreateMapperForFormattedTemplate();

        Stream ITemplateFormatter.ToFormattedTemplate(Model.ProvisioningTemplate template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            // AutoMapper configuration
            var config = new MapperConfiguration(cfg =>
            {
                cfg.AddProfile(new V201605Profile());
            });

            // Configure and use AutoMapper
            var mapper = this.CreateMapperForFormattedTemplate();
            var result = mapper.Map<V201605.ProvisioningTemplate>(template);

            V201605.Provisioning wrappedResult = new V201605.Provisioning();
            wrappedResult.Preferences = new V201605.Preferences
            {
                Generator = this.GetType().Assembly.FullName
            };
            wrappedResult.Templates = new V201605.Templates[] {
                new V201605.Templates
                {
                    ID = $"CONTAINER-{template.Id}",
                    ProvisioningTemplate = new V201605.ProvisioningTemplate[]
                    {
                        result
                    }
                }
            };

            XmlSerializerNamespaces ns =
                new XmlSerializerNamespaces();
            ns.Add(((IXMLSchemaFormatter)this).NamespacePrefix,
                ((IXMLSchemaFormatter)this).NamespaceUri);

            var output = XMLSerializer.SerializeToStream<V201605.Provisioning>(wrappedResult, ns);
            output.Position = 0;
            return (output);
        }

        /// <summary>
        /// Abstract method used to prepare a Mapper object ready to map from XML object model to the Domain Model 
        /// </summary>
        /// <returns></returns>
        protected abstract IMapper CreateMapperForProvisioningTemplate();

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
            V201605.ProvisioningTemplate source = null;

            // Prepare a variable to hold the resulting ProvisioningTemplate instance
            Model.ProvisioningTemplate result = new Model.ProvisioningTemplate();

            // Configure AutoMapper
            var mapper = this.CreateMapperForProvisioningTemplate();

            // Determine if we're working on a wrapped SharePointProvisioningTemplate or not
            if (xml.Root.Name == pnp + "Provisioning")
            {
                // Deserialize the whole wrapper
                V201605.Provisioning wrappedResult = XMLSerializer.Deserialize<V201605.Provisioning>(xml);

                // Handle the wrapper schema parameters
                if (wrappedResult.Preferences != null &&
                    wrappedResult.Preferences.Parameters != null &&
                    wrappedResult.Preferences.Parameters.Length > 0)
                {
                    foreach (var parameter in wrappedResult.Preferences.Parameters)
                    {
                        result.Parameters.Add(parameter.Key, parameter.Text != null ? parameter.Text.Aggregate(String.Empty, (acc, i) => acc + i) : null);
                    }
                }

                // Handle Localizations
                if (wrappedResult.Localizations != null)
                {
                    result.Localizations.AddRange(
                        from l in wrappedResult.Localizations
                        select new Localization
                        {
                            LCID = l.LCID,
                            Name = l.Name,
                            ResourceFile = l.ResourceFile,
                        });
                }

                foreach (var templates in wrappedResult.Templates)
                {
                    // Let's see if we have an in-place template with the provided ID or if we don't have a provided ID at all
                    source = templates.ProvisioningTemplate.FirstOrDefault(spt => spt.ID == identifier || String.IsNullOrEmpty(identifier));

                    // If we don't have a template, but there are external file references
                    if (source == null && templates.ProvisioningTemplateFile.Length > 0)
                    {
                        // Otherwise let's see if we have an external file for the template
                        var externalSource = templates.ProvisioningTemplateFile.FirstOrDefault(sptf => sptf.ID == identifier);

                        Stream externalFileStream = this._provider.Connector.GetFileStream(externalSource.File);
                        xml = XDocument.Load(externalFileStream);

                        if (xml.Root.Name != pnp + "ProvisioningTemplate")
                        {
                            throw new ApplicationException("Invalid external file format. Expected a ProvisioningTemplate file!");
                        }
                        else
                        {
                            source = XMLSerializer.Deserialize<V201605.ProvisioningTemplate>(xml);
                        }
                    }

                    if (source != null)
                    {
                        break;
                    }
                }
            }
            else if (xml.Root.Name == pnp + "ProvisioningTemplate")
            {
                var IdAttribute = xml.Root.Attribute("ID");

                // If there is a provided ID, and if it doesn't equal the current ID
                if (!String.IsNullOrEmpty(identifier) &&
                    IdAttribute != null &&
                    IdAttribute.Value != identifier)
                {
                    // TODO: Use resource file
                    throw new ApplicationException("The provided template identifier is not available!");
                }
                else
                {
                    source = XMLSerializer.Deserialize<V201605.ProvisioningTemplate>(xml);
                }
            }

            // Apply AutoMapping
            result = mapper.Map<Model.ProvisioningTemplate>(source);

            return (result);
        }
    }
}

