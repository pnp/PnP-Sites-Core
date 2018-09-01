using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System.Xml.XPath;
using System.Xml;
using System.Xml.Serialization;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Helper class that abstracts from any specific version of XMLPnPSchemaFormatter
    /// </summary>
    public class XMLProvisioningHierarchyFormatter : IProvisioningHierarchyFormatter
    {
        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        public bool IsValid(Stream hierarchy)
        {
            XMLPnPSchemaFormatter innerFormatter = new XMLPnPSchemaFormatter();
            return (innerFormatter.IsValid(hierarchy));
        }

        public Stream ToFormattedHierarchy(ProvisioningHierarchy hierarchy)
        {
            throw new NotImplementedException();
        }

        public ProvisioningHierarchy ToProvisioningHierarchy(Stream hierarchy)
        {
            // Create a copy of the source stream
            MemoryStream sourceStream = new MemoryStream();
            hierarchy.Position = 0;
            hierarchy.CopyTo(sourceStream);
            sourceStream.Position = 0;

            // Check the provided template against the XML schema
            if (!this.IsValid(sourceStream))
            {
                // TODO: Use resource file
                throw new ApplicationException("The provided provisioning file is not valid!");
            }

            // Prepare the output variable
            ProvisioningHierarchy resultHierarchy = new ProvisioningHierarchy();

            // Determine if the file is a provisioning hierarchy
            sourceStream.Position = 0;
            XDocument xml = XDocument.Load(sourceStream);
            if (xml.Root.Name.LocalName != "Provisioning")
            {
                throw new ApplicationException("The provided provisioning file is not a Hierarchy!");
            }

            // Determine the specific formatter needed for the current provisioning file
            var innerFormatter = XMLPnPSchemaFormatter.GetSpecificFormatter(
                xml.Root.Name.NamespaceName);

            // Process all the provisioning templates included in the hierarchy, if any
            XmlNamespaceManager nsManager = new XmlNamespaceManager(new System.Xml.NameTable());
            nsManager.AddNamespace("pnp", xml.Root.Name.NamespaceName);

            // Start with templates embedded in the provisioning file
            var templates = xml.XPathSelectElements("/pnp:Provisioning/pnp:Templates/pnp:ProvisioningTemplate", nsManager).ToList();

            foreach (var template in templates)
            {
                // Save the single template into a MemoryStream
                MemoryStream templateStream = new MemoryStream();
                template.Save(templateStream);
                templateStream.Position = 0;

                // Process the single template with the classic technique
                var provisioningTemplate = innerFormatter.ToProvisioningTemplate(templateStream);

                // Add the generated template to the resulting hierarchy
                resultHierarchy.Templates.Add(provisioningTemplate);
            }

            // Then process any external file reference
            var templateFiles = xml.XPathSelectElements("/pnp:Provisioning/pnp:Templates/pnp:ProvisioningTemplateFile", nsManager).ToList();

            foreach (var template in templateFiles)
            {
                var templateID = template.Attribute("ID")?.Value;
                var templateFile = template.Attribute("File")?.Value;
                if (!String.IsNullOrEmpty(templateFile) && !String.IsNullOrEmpty(templateID))
                {
                    // Process the single template file with the classic technique
                    var provisioningTemplate = this._provider.GetTemplate(templateFile);
                    provisioningTemplate.Id = templateID;

                    // Add the generated template to the resulting hierarchy
                    resultHierarchy.Templates.Add(provisioningTemplate);
                }
            }

            // And now process the top level children elements
            // using schema specific serializers
            // TODO: Preferences, Localization, Tenant, Sequences

            // We prepare a dummy template to leverage the existing serialization infrastructure
            var dummyTemplate = new ProvisioningTemplate();

            // Deserialize the whole wrapper
            Object wrapper = null;
            var wrapperType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Provisioning, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
            XmlSerializer xmlSerializer = new XmlSerializer(wrapperType);
            using (var reader = xml.Root.CreateReader())
            {
                wrapper = xmlSerializer.Deserialize(reader);
            }

            // Handle the Parameters of the schema wrapper, if any
            var tps = new TemplateParametersSerializer();
            tps.Deserialize(wrapper, dummyTemplate);

            // Handle the Localizations of the schema wrapper, if any
            var ls = new LocalizationsSerializer();
            ls.Deserialize(wrapper, dummyTemplate);

            // Handle the Tenant-wide settings of the schema wrapper, if any
            var ts = new TenantSerializer();
            ts.Deserialize(wrapper, dummyTemplate);

            return (resultHierarchy);
        }
    }
}
