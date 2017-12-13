using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System.Xml.Serialization;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers;
using System.Collections;
using System.Reflection;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    internal abstract class XmlPnPSchemaBaseSerializer<TSchemaTemplate> : IXMLSchemaFormatter, ITemplateFormatter
        where TSchemaTemplate: new()
    {
        private TemplateProviderBase _provider;
        private Stream _referenceSchema;

        protected TemplateProviderBase Provider => _provider;

        public XmlPnPSchemaBaseSerializer(Stream referenceSchema)
        {
            if (referenceSchema == null)
            {
                throw new ArgumentNullException("referenceSchema");
            }

            this._referenceSchema = referenceSchema;
        }

        public abstract string NamespacePrefix { get; }
        public abstract string NamespaceUri { get; }

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
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
            this._referenceSchema.Seek(0, SeekOrigin.Begin);
            schemas.Add(((IXMLSchemaFormatter)this).NamespaceUri,
                new XmlTextReader(this._referenceSchema));
                        
            Boolean result = true;
            xml.Validate(schemas, (o, e) =>
            {
                Diagnostics.Log.Error(e.Exception, "SchemaFormatter", "Template is not valid: {0}", e.Message);
                result = false;
            });

            return (result);
        }

        protected Object ProcessInputStream(Stream template, string identifier, ProvisioningTemplate result)
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
            XNamespace pnp = this.NamespaceUri;

            // Prepare a variable to hold the single source formatted template
            TSchemaTemplate source = default(TSchemaTemplate);

            // Determine if we're working on a wrapped ProvisioningTemplate or not
            if (xml.Root.Name == pnp + "Provisioning")
            {
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
                tps.Deserialize(wrapper, result);

                // Handle the Localizations of the schema wrapper, if any
                var ls = new LocalizationsSerializer();
                ls.Deserialize(wrapper, result);

                // Handle the Tenant-wide settings of the schema wrapper, if any
                var ts = new TenantSerializer();
                ts.Deserialize(wrapper, result);

                // Get the list of templates, if any, wrapped by the wrapper
                var wrapperTemplates = wrapperType.GetProperty("Templates", 
                    System.Reflection.BindingFlags.Instance | 
                    System.Reflection.BindingFlags.Public | 
                    System.Reflection.BindingFlags.IgnoreCase).GetValue(wrapper);

                if (wrapperTemplates != null)
                {
                    // Search for the requested Provisioning Template
                    foreach (var templates in (IEnumerable)wrapperTemplates)
                    {
                        // Let's see if we have an in-place template with the provided ID or if we don't have a provided ID at all
                        var provisioningTemplates = templates.GetType()
                            .GetProperty("ProvisioningTemplate",
                                System.Reflection.BindingFlags.Instance |
                                System.Reflection.BindingFlags.Public |
                                System.Reflection.BindingFlags.IgnoreCase).GetValue(templates);

                        if (provisioningTemplates != null)
                        {
                            foreach (var t in (IEnumerable)provisioningTemplates)
                            {
                                var templateId = (String)t.GetType().GetProperty("ID",
                                    System.Reflection.BindingFlags.Instance |
                                    System.Reflection.BindingFlags.Public |
                                    System.Reflection.BindingFlags.IgnoreCase).GetValue(t);

                                if ((templateId != null && templateId == identifier) || String.IsNullOrEmpty(identifier))
                                {
                                    source = (TSchemaTemplate)t;
                                }
                            }

                            if (source == null)
                            {
                                var provisioningTemplateFiles = templates.GetType()
                                    .GetProperty("ProvisioningTemplateFile",
                                        System.Reflection.BindingFlags.Instance |
                                        System.Reflection.BindingFlags.Public |
                                        System.Reflection.BindingFlags.IgnoreCase).GetValue(templates);

                                // If we don't have a template, but there are external file references
                                if (source == null && provisioningTemplateFiles != null)
                                {
                                    foreach (var f in (IEnumerable)provisioningTemplateFiles)
                                    {
                                        var templateId = (String)f.GetType().GetProperty("ID",
                                            System.Reflection.BindingFlags.Instance |
                                            System.Reflection.BindingFlags.Public |
                                            System.Reflection.BindingFlags.IgnoreCase).GetValue(f);

                                        if ((templateId != null && templateId == identifier) || String.IsNullOrEmpty(identifier))
                                        {
                                            // Let's see if we have an external file for the template
                                            var externalFile = (String)f.GetType().GetProperty("File",
                                                System.Reflection.BindingFlags.Instance |
                                                System.Reflection.BindingFlags.Public |
                                                System.Reflection.BindingFlags.IgnoreCase).GetValue(f);

                                            Stream externalFileStream = this.Provider.Connector.GetFileStream(externalFile);
                                            xml = XDocument.Load(externalFileStream);

                                            if (xml.Root.Name != pnp + "ProvisioningTemplate")
                                            {
                                                throw new ApplicationException("Invalid external file format. Expected a ProvisioningTemplate file!");
                                            }
                                            else
                                            {
                                                source = XMLSerializer.Deserialize<TSchemaTemplate>(xml);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (source != null)
                        {
                            break;
                        }
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
                    source = XMLSerializer.Deserialize<TSchemaTemplate>(xml);
                }
            }

            return (source);
        }

        public ProvisioningTemplate ToProvisioningTemplate(Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        public ProvisioningTemplate ToProvisioningTemplate(Stream template, string identifier)
        {
            using (var scope = new PnPSerializationScope(typeof(TSchemaTemplate)))
            {
                // Prepare a variable to hold the resulting ProvisioningTemplate instance
                var result = new ProvisioningTemplate();

                // Prepare a variable to hold the single source formatted template
                var source = ProcessInputStream(template, identifier, result);

                DeserializeTemplate(source, result);

                return (result);
            }
        }

        protected virtual void DeserializeTemplate(Object persistenceTemplate, ProvisioningTemplate template)
        {
            // Get all serializers to run in automated mode, ordered by DeserializationSequence
            var currentAssembly = this.GetType().Assembly;

            XMLPnPSchemaVersion currentSchemaVersion = GetCurrentSchemaVersion();

            var serializers = currentAssembly.GetTypes()
                .Where(t => t.GetInterface(typeof(IPnPSchemaSerializer).FullName) != null
                       && t.BaseType.Name == typeof(Xml.PnPBaseSchemaSerializer<>).Name)
                .Where(t => 
                {
                    var a = t.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                    return (a.MinimalSupportedSchemaVersion <= currentSchemaVersion && a.DeserializationSequence >= 0);
                })
                .OrderByDescending(s =>
                {
                    var a = s.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                    return (a.MinimalSupportedSchemaVersion);
                }
                )
                .GroupBy(t => t.BaseType.GenericTypeArguments.FirstOrDefault()?.FullName)
                .OrderBy(g =>
                {
                    var maxInGroup = g.OrderByDescending(s =>
                    {
                        var a = s.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                        return (a.MinimalSupportedSchemaVersion);
                    }
                    ).FirstOrDefault();
                    return (maxInGroup.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault()?.SerializationSequence);
                });

            foreach (var group in serializers)
            {
                var serializerType = group.FirstOrDefault();
                if (serializerType != null)
                {
                    var serializer = Activator.CreateInstance(serializerType) as IPnPSchemaSerializer;
                    if (serializer != null)
                    {
                        serializer.Deserialize(persistenceTemplate, template);
                    }
                }
            }
        }

        public Stream ToFormattedTemplate(ProvisioningTemplate template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            using (var scope = new PnPSerializationScope(typeof(TSchemaTemplate)))
            {
                var result = new TSchemaTemplate();

                // Create the wrapper
                var wrapperType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Provisioning, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                Object wrapper = Activator.CreateInstance(wrapperType);

                // Create the Preferences
                var preferencesType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Preferences, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                Object preferences = Activator.CreateInstance(preferencesType);

                wrapper.GetType().GetProperty("Preferences",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase).SetValue(wrapper, preferences);

                // Handle the Parameters of the schema wrapper, if any
                var tps = new TemplateParametersSerializer();
                tps.Serialize(template, wrapper);

                // Handle the Localizations of the schema wrapper, if any
                var ls = new LocalizationsSerializer();
                ls.Serialize(template, wrapper);

                // Handle the Tenant-wide of the schema wrapper, if any
                var ts = new TenantSerializer();
                ts.Serialize(template, wrapper);

                // Configure the Generator
                preferences.GetType().GetProperty("Generator",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase).SetValue(preferences, this.GetType().Assembly.FullName);

                // Configure the output Template
                var templatesType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Templates, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var templates = Array.CreateInstance(templatesType, 1);
                var templatesItem = Activator.CreateInstance(templatesType);

                templatesItem.GetType().GetProperty("ID",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase).SetValue(templatesItem, $"CONTAINER-{template.Id}");

                var provisioningTemplates = Array.CreateInstance(typeof(TSchemaTemplate), 1);
                provisioningTemplates.SetValue(result, 0);

                templatesItem.GetType().GetProperty("ProvisioningTemplate",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase).SetValue(templatesItem, provisioningTemplates);

                templates.SetValue(templatesItem, 0);

                wrapperType.GetProperty("Templates",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase).SetValue(wrapper, templates);

                SerializeTemplate(template, result);

                XmlSerializerNamespaces ns =
                    new XmlSerializerNamespaces();
                ns.Add(((IXMLSchemaFormatter)this).NamespacePrefix,
                    ((IXMLSchemaFormatter)this).NamespaceUri);

                MemoryStream output = new MemoryStream();
                XmlSerializer xmlSerializer = new XmlSerializer(wrapperType);
                if (ns != null)
                {
                    xmlSerializer.Serialize(output, wrapper, ns);
                }
                else
                {
                    xmlSerializer.Serialize(output, wrapper);
                }

                output.Position = 0;
                return (output);
            }
        }

        protected virtual void SerializeTemplate(ProvisioningTemplate template, Object persistenceTemplate)
        {
            // Get all serializers to run in automated mode, ordered by DeserializationSequence
            var currentAssembly = this.GetType().Assembly;

            XMLPnPSchemaVersion currentSchemaVersion = GetCurrentSchemaVersion();

            var serializers = currentAssembly.GetTypes()
                .Where(t => t.GetInterface(typeof(IPnPSchemaSerializer).FullName) != null
                       && t.BaseType.Name == typeof(Xml.PnPBaseSchemaSerializer<>).Name)
                .Where(t =>
                {
                    var a = t.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                    return (a.MinimalSupportedSchemaVersion <= currentSchemaVersion && a.SerializationSequence >= 0);
                })
                .OrderByDescending(s =>
                {
                    var a = s.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                    return (a.MinimalSupportedSchemaVersion);
                }
                )
                .GroupBy(t => t.BaseType.GenericTypeArguments.FirstOrDefault()?.FullName)
                .OrderBy(g =>
                {
                    var maxInGroup = g.OrderByDescending(s =>
                    {
                        var a = s.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                        return (a.MinimalSupportedSchemaVersion);
                    }
                    ).FirstOrDefault();
                    return (maxInGroup.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault()?.DeserializationSequence);
                });

            foreach (var group in serializers)
            {
                var serializerType = group.FirstOrDefault();
                if (serializerType != null)
                {
                    var serializer = Activator.CreateInstance(serializerType) as IPnPSchemaSerializer;
                    if (serializer != null)
                    {
                        serializer.Serialize(template, persistenceTemplate);
                    }
                }
            }
        }

        private static XMLPnPSchemaVersion GetCurrentSchemaVersion()
        {
            var currentSchemaTemplateNamespace = typeof(TSchemaTemplate).Namespace;
            var currentSchemaVersionString = $"V{currentSchemaTemplateNamespace.Substring(currentSchemaTemplateNamespace.IndexOf(".Xml.") + 6)}";
            var currentSchemaVersion = (XMLPnPSchemaVersion)Enum.Parse(typeof(XMLPnPSchemaVersion), currentSchemaVersionString);
            return currentSchemaVersion;
        }
    }
}
