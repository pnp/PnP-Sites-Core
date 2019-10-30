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
using System.Xml.XPath;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Base class for serialization/deserialization of provisioning templates
    /// with the new schema serializer
    /// </summary>
    /// <typeparam name="TSchemaTemplate"></typeparam>
    internal abstract class XmlPnPSchemaBaseSerializer<TSchemaTemplate> : IXMLSchemaFormatter, ITemplateFormatterWithValidation, IProvisioningHierarchyFormatter
        where TSchemaTemplate: new()
    {
        private TemplateProviderBase _provider;
        private Stream _referenceSchema;

        protected TemplateProviderBase Provider => _provider;

        public XmlPnPSchemaBaseSerializer(Stream referenceSchema)
        {
            this._referenceSchema = referenceSchema ?? 
                throw new ArgumentNullException("referenceSchema");
        }

        public abstract string NamespacePrefix { get; }
        public abstract string NamespaceUri { get; }

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        /// <summary>
        /// Checks if the provided source Stream (the XML) is valid against the current XSD schema
        /// </summary>
        /// <param name="template">The source Stream (the XML)</param>
        /// <returns>Whether the XML template is valid or not</returns>
        public bool IsValid(Stream template)
        {
            return GetValidationResults(template).IsValid;
        }

        /// <summary>
        /// Checks if the provided source Stream (the XML) is valid against the current XSD schema
        /// </summary>
        /// <param name="template">The source Stream (the XML)</param>
        /// <returns>Whether the XML template is valid or not</returns>
        public ValidationResult GetValidationResults(Stream template)
        {
            var exceptions = new List<Exception>();
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
                exceptions.Add(e.Exception);
                Diagnostics.Log.Error(e.Exception, "SchemaFormatter", "Template is not valid: {0}", e.Message);
                result = false;
            });

            return new ValidationResult { IsValid = result, Exceptions = exceptions };
        }

        //public bool IsValid(Stream template, Action<Exception,string> validationErrorMessageAction)
        //{
        //    if(template == null)
        //    {
        //        throw new ArgumentNullException(nameof(template));
        //    }
        //    if(validationErrorMessageAction == null)
        //    {
        //        throw new ArgumentNullException(nameof(validationErrorMessageAction));
        //    }

        //    // Load the template into an XDocument
        //    var xml = XDocument.Load(template);

        //    // Prepare the XML Schema Set
        //    var schemas = new XmlSchemaSet();
        //    this._referenceSchema.Seek(0, SeekOrigin.Begin);
        //    schemas.Add(((IXMLSchemaFormatter)this).NamespaceUri, new XmlTextReader(this._referenceSchema));

        //    bool result = true;
        //    xml.Validate(schemas, (o, e) =>
        //    {
        //        Diagnostics.Log.Error(e.Exception, "SchemaFormatter", "Template is not valid: {0}", e.Message);
        //        validationErrorMessageAction(e.Exception, e.Message);
        //        result = false;
        //    });
        //    return result;
        //}
        /// <summary>
        /// Converts a Stream of bytes (the XML) into a XML-based object created using XmlSerializer
        /// </summary>
        /// <param name="template">The source Stream of bytes (the XML)</param>
        /// <param name="identifier">An optional identifier for the template to extract from the XML</param>
        /// <param name="result">A reference ProvisioningTemplate object</param>
        /// <returns>The resulting XML-based object extracted from the Stream</returns>
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
            var validationResult = this.GetValidationResults(sourceStream);
            if (!validationResult.IsValid)
            {
                throw new ApplicationException("Template is not valid", new AggregateException(validationResult.Exceptions));
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

                // Get all Provisioning-level serializers to run in automated mode, ordered by DeserializationSequence
                var serializers = GetSerializersForCurrentContext(SerializerScope.Provisioning, a => a?.DeserializationSequence);

                // Invoke all the Provisioning-level serializers
                InvokeSerializers(result, wrapper, serializers, SerializationAction.Deserialize);

                // Get the list of templates, if any, wrapped by the wrapper
                var wrapperTemplates = wrapper.GetPublicInstancePropertyValue("Templates");

                if (wrapperTemplates != null)
                {
                    // Search for the requested Provisioning Template
                    foreach (var templates in (IEnumerable)wrapperTemplates)
                    {
                        // Let's see if we have an in-place template with the provided ID or if we don't have a provided ID at all
                        var provisioningTemplates = templates.GetPublicInstancePropertyValue("ProvisioningTemplate");

                        if (provisioningTemplates != null)
                        {
                            foreach (var t in (IEnumerable)provisioningTemplates)
                            {
                                var templateId = t.GetPublicInstancePropertyValue("ID") as String;

                                if ((templateId != null && templateId == identifier) || String.IsNullOrEmpty(identifier))
                                {
                                    source = (TSchemaTemplate)t;
                                }
                            }

                            if (source == null)
                            {
                                var provisioningTemplateFiles = templates.GetPublicInstancePropertyValue("ProvisioningTemplateFile");

                                // If we don't have a template, but there are external file references
                                if (source == null && provisioningTemplateFiles != null)
                                {
                                    foreach (var f in (IEnumerable)provisioningTemplateFiles)
                                    {
                                        var templateId = f.GetPublicInstancePropertyValue("ID") as String;

                                        if ((templateId != null && templateId == identifier) || String.IsNullOrEmpty(identifier))
                                        {
                                            // Let's see if we have an external file for the template
                                            var externalFile = f.GetPublicInstancePropertyValue("File") as String;

                                            if (!String.IsNullOrEmpty(externalFile))
                                            {
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

        /// <summary>
        /// Deserializes a Stream of bytes (the XML) into a Provisioning Template
        /// </summary>
        /// <param name="template">The source Stream of bytes (the XML)</param>
        /// <returns>The deserialized Provisioning Template</returns>
        public ProvisioningTemplate ToProvisioningTemplate(Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        /// <summary>
        /// Deserializes a Stream of bytes (the XML) into a Provisioning Template, based on an optional identifier
        /// </summary>
        /// <param name="template">The source Stream of bytes (the XML)</param>
        /// <param name="identifier">An optional identifier for the template to deserialize</param>
        /// <returns>The deserialized Provisioning Template</returns>
        public ProvisioningTemplate ToProvisioningTemplate(Stream template, string identifier)
        {
            using (var scope = new PnPSerializationScope(typeof(TSchemaTemplate)))
            {
                // Prepare a variable to hold the resulting ProvisioningTemplate instance
                var result = new ProvisioningTemplate();

                // Prepare a variable to hold the single source formatted template
                // We provide the result instance of ProvisioningTemplate in order
                // to configure the tenant/hierarchy level items
                // We get back the XML-based object to use with the other serializers
                var source = ProcessInputStream(template, identifier, result);

                // We process the chain of deserialization 
                // with the Provisioning-level serializers
                DeserializeTemplate(source, result);

                return (result);
            }
        }

        /// <summary>
        /// This method deserializes an XML-based object, created with XmlSerializer, into a Provisioning Template
        /// </summary>
        /// <param name="persistenceTemplate">The XML-based object</param>
        /// <param name="template">The resulting template</param>
        protected virtual void DeserializeTemplate(Object persistenceTemplate, ProvisioningTemplate template)
        {
            // Get all ProvisioningTemplate-level serializers to run in automated mode, ordered by DeserializationSequence
            var serializers = GetSerializersForCurrentContext(SerializerScope.ProvisioningTemplate, a => a?.DeserializationSequence);

            // Invoke all the ProvisioningTemplate-level serializers
            InvokeSerializers(template, persistenceTemplate, serializers, SerializationAction.Deserialize);
        }

        /// <summary>
        /// Serializes an in-memory ProvisioningTemplate into a Stream (the XML)
        /// </summary>
        /// <param name="template">The ProvisioningTemplate to serialize</param>
        /// <returns>The resulting Stream (the XML)</returns>
        public Stream ToFormattedTemplate(ProvisioningTemplate template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            using (var scope = new PnPSerializationScope(typeof(TSchemaTemplate)))
            {
                var result = new TSchemaTemplate();
                Stream output = null;

                // Process the template to generate the output stream
                output = ProcessOutputStream(template, result);

                return (output);
            }
        }

        /// <summary>
        /// Serializes an in-memory ProvisioningTemplate into a Stream (the XML)
        /// </summary>
        /// <param name="template">The ProvisioningTemplate to serialize</param>
        /// <param name="result">The typed XML-based object defined using XmlSerializer</param>
        /// <returns>The resulting Stream (the XML)</returns>
        protected Stream ProcessOutputStream(ProvisioningTemplate template, TSchemaTemplate result)
        {
            // Prepare the output wrapper
            Type wrapperType;
            object wrapper, templatesItem;
            Array templates;

            // Process the hierarchy part of the template
            ProcessOutputHierarchy(template, out wrapperType, out wrapper, out templates, out templatesItem);

            // Add the single template to the output
            var provisioningTemplates = Array.CreateInstance(typeof(TSchemaTemplate), 1);
            provisioningTemplates.SetValue(result, 0);

            templatesItem.SetPublicInstancePropertyValue("ProvisioningTemplate", provisioningTemplates);

            templates.SetValue(templatesItem, 0);

            wrapper.SetPublicInstancePropertyValue("Templates", templates);

            // Serialize the template mapping the ProvisioningTemplate object to the XML-based object
            SerializeTemplate(template, result);

            // Serialize the XML-based object into a Stream (the XML)
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

            // Re-base the Stream and return it
            output.Position = 0;
            return (output);
        }

        /// <summary>
        /// Prepares a ProvisioningTemplate to be wrapped into the Hierarchy container object
        /// </summary>
        /// <param name="template">The ProvisioningTemplate to wrap</param>
        /// <param name="wrapperType">The Type of the wrapper</param>
        /// <param name="wrapper">The wrapper</param>
        /// <param name="templates">The collection of template within the wrapper</param>
        /// <param name="templatesItem">The template to add</param>
        private void ProcessOutputHierarchy(ProvisioningTemplate template, out Type wrapperType, out object wrapper, out Array templates, out object templatesItem)
        {
            // Create the wrapper
            wrapperType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Provisioning, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
            wrapper = Activator.CreateInstance(wrapperType);

            // Create the Preferences
            var preferencesType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Preferences, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
            Object preferences = Activator.CreateInstance(preferencesType);

            wrapper.SetPublicInstancePropertyValue("Preferences", preferences);

            // Get all Provisioning-level serializers to run in automated mode, ordered by SerializationSequence
            var serializers = GetSerializersForCurrentContext(SerializerScope.Provisioning, a => a?.SerializationSequence);

            // Invoke all the Provisioning-level serializers
            InvokeSerializers(template, wrapper, serializers, SerializationAction.Serialize);

            // Get all Tenant-levelserializers to run in automated mode, ordered by SerializationSequence
            serializers = GetSerializersForCurrentContext(SerializerScope.Tenant, a => a?.SerializationSequence);

            // Invoke all the Tenant-levelserializers
            InvokeSerializers(template, wrapper, serializers, SerializationAction.Serialize);

            // Configure the basic properties of the wrapper
            if (template.ParentHierarchy != null)
            {
                wrapper.SetPublicInstancePropertyValue("Author", template.ParentHierarchy.Author);
                wrapper.SetPublicInstancePropertyValue("DisplayName", template.ParentHierarchy.DisplayName);
                wrapper.SetPublicInstancePropertyValue("Description", template.ParentHierarchy.Description);
                wrapper.SetPublicInstancePropertyValue("ImagePreviewUrl", template.ParentHierarchy.ImagePreviewUrl);
                wrapper.SetPublicInstancePropertyValue("Generator", template.ParentHierarchy.Generator);
                wrapper.SetPublicInstancePropertyValue("Version", (Decimal)template.ParentHierarchy.Version);
            }

            // Configure the Generator
            preferences.SetPublicInstancePropertyValue("Generator", this.GetType().Assembly.FullName);

            // Configure the output Template
            var templatesType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Templates, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
            templates = Array.CreateInstance(templatesType, 1);
            templatesItem = Activator.CreateInstance(templatesType);
            templatesItem.SetPublicInstancePropertyValue("ID", $"CONTAINER-{template.Id}");
        }

        /// <summary>
        /// Serializes a ProvisioningTemplate into a XML-based object generated with XmlSerializer
        /// </summary>
        /// <param name="template">The ProvisioningTemplate to serialize</param>
        /// <param name="persistenceTemplate">The XML-based object to serialize the template into</param>
        protected virtual void SerializeTemplate(ProvisioningTemplate template, Object persistenceTemplate)
        {
            // Get all ProvisioningTemplate-level serializers to run in automated mode, ordered by DeserializationSequence
            var serializers = GetSerializersForCurrentContext(SerializerScope.ProvisioningTemplate, a => a?.SerializationSequence);

            // Invoke all the ProvisioningTemplate-level serializers
            InvokeSerializers(template, persistenceTemplate, serializers, SerializationAction.Serialize);
        }

        /// <summary>
        /// Allows to retrieve the current XML Schema version
        /// </summary>
        /// <returns>The current XML schema version</returns>
        private static XMLPnPSchemaVersion GetCurrentSchemaVersion()
        {
            var currentSchemaTemplateNamespace = typeof(TSchemaTemplate).Namespace;
            var currentSchemaVersionString = $"V{currentSchemaTemplateNamespace.Substring(currentSchemaTemplateNamespace.IndexOf(".Xml.") + 6)}";
            var currentSchemaVersion = (XMLPnPSchemaVersion)Enum.Parse(typeof(XMLPnPSchemaVersion), currentSchemaVersionString);
            return currentSchemaVersion;
        }

        /// <summary>
        /// Serializes a ProvisioningHierarchy into a Stream (the XML)
        /// </summary>
        /// <param name="hierarchy">The ProvisioningHierarchy to serialize</param>
        /// <returns>The resulting Stream (the XML)</returns>
        public Stream ToFormattedHierarchy(ProvisioningHierarchy hierarchy)
        {
            if (hierarchy == null)
            {
                throw new ArgumentNullException(nameof(hierarchy));
            }

            using (var scope = new PnPSerializationScope(typeof(TSchemaTemplate)))
            {
                // We prepare a dummy template to leverage the existing deserialization infrastructure
                var dummyTemplate = new ProvisioningTemplate();
                dummyTemplate.Id = $"DUMMY-{Guid.NewGuid()}";
                hierarchy.Templates.Add(dummyTemplate);

                // Prepare the output wrapper
                Type wrapperType;
                object wrapper, templatesItem;
                Array templates;

                ProcessOutputHierarchy(dummyTemplate, out wrapperType, out wrapper, out templates, out templatesItem);

                // Handle the Sequences, if any
                // Get all ProvisioningHierarchy-level serializers to run in automated mode, ordered by SerializationSequence
                var serializers = GetSerializersForCurrentContext(SerializerScope.ProvisioningHierarchy, a => a?.SerializationSequence);

                // Invoke all the ProvisioningHierarchy-level serializers
                InvokeSerializers(dummyTemplate, wrapper, serializers, SerializationAction.Serialize);

                // Remove the dummy template
                hierarchy.Templates.Remove(dummyTemplate);

                // Add every single template to the output
                var provisioningTemplates = Array.CreateInstance(typeof(TSchemaTemplate), hierarchy.Templates.Count);
                for (int c = 0; c < hierarchy.Templates.Count; c++)
                {
                    // Prepare variable to hold the output template
                    var outputTemplate = new TSchemaTemplate();

                    // Serialize the real templates
                    SerializeTemplate(hierarchy.Templates[c], outputTemplate);

                    // Add the serialized template to the output
                    provisioningTemplates.SetValue(outputTemplate, c);
                }

                templatesItem.SetPublicInstancePropertyValue("ProvisioningTemplate", provisioningTemplates);

                templates.SetValue(templatesItem, 0);

                wrapper.SetPublicInstancePropertyValue("Templates", templates);

                // Serialize the XML-based object into a Stream
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

                // Re-base the Stream and return it
                output.Position = 0;
                return (output);
            }
        }

        /// <summary>
        /// Deserializes a source Stream (the XML) into a ProvisioningHierarchy 
        /// </summary>
        /// <param name="hierarchy">The source Stream (the XML)</param>
        /// <returns>The resulting ProvisioningHierarchy object</returns>
        public ProvisioningHierarchy ToProvisioningHierarchy(Stream hierarchy)
        {
            // Create a copy of the source stream
            MemoryStream sourceStream = new MemoryStream();
            hierarchy.Position = 0;
            hierarchy.CopyTo(sourceStream);
            sourceStream.Position = 0;

            // Check the provided template against the XML schema
            var validationResult = this.GetValidationResults(sourceStream);
            if (!validationResult.IsValid)
            {
                // TODO: Use resource file
                throw new ApplicationException("Template is not valid", new AggregateException(validationResult.Exceptions));
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

            using (var scope = new PnPSerializationScope(typeof(TSchemaTemplate)))
            {
                // We prepare a dummy template to leverage the existing serialization infrastructure
                var dummyTemplate = new ProvisioningTemplate();
                dummyTemplate.Id = $"DUMMY-{Guid.NewGuid()}";
                resultHierarchy.Templates.Add(dummyTemplate);

                // Deserialize the whole wrapper
                Object wrapper = null;
                var wrapperType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Provisioning, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                XmlSerializer xmlSerializer = new XmlSerializer(wrapperType);
                using (var reader = xml.Root.CreateReader())
                {
                    wrapper = xmlSerializer.Deserialize(reader);
                }

                #region Process Provisioning level serializers

                // Get all serializers to run in automated mode, ordered by DeserializationSequence
                var serializers = GetSerializersForCurrentContext(SerializerScope.Provisioning, a => a?.DeserializationSequence);

                // Invoke all the serializers
                InvokeSerializers(dummyTemplate, wrapper, serializers, SerializationAction.Deserialize);

                #endregion

                #region Process Tenant level serializers

                // Get all serializers to run in automated mode, ordered by DeserializationSequence
                serializers = GetSerializersForCurrentContext(SerializerScope.Tenant, a => a?.DeserializationSequence);

                // Invoke all the serializers
                InvokeSerializers(dummyTemplate, wrapper, serializers, SerializationAction.Deserialize);

                #endregion

                #region Process ProvisioningHierarchy level serializers

                // Get all serializers to run in automated mode, ordered by DeserializationSequence
                serializers = GetSerializersForCurrentContext(SerializerScope.ProvisioningHierarchy, a => a?.DeserializationSequence);

                // Invoke all the serializers
                InvokeSerializers(dummyTemplate, wrapper, serializers, SerializationAction.Deserialize);

                #endregion

                // Remove the dummy template from the hierarchy
                resultHierarchy.Templates.Remove(dummyTemplate);
            }

            return (resultHierarchy);
        }

        private IOrderedEnumerable<IGrouping<string, Type>> GetSerializersForCurrentContext(SerializerScope scope,
            Func<TemplateSchemaSerializerAttribute, Int32?> sortingSelector)
        {
            // Get all serializers to run in automated mode, ordered by sortingSelector
            var currentAssembly = this.GetType().Assembly;

            XMLPnPSchemaVersion currentSchemaVersion = GetCurrentSchemaVersion();

            var serializers = currentAssembly.GetTypes()
                // Get all the serializers
                .Where(t => t.GetInterface(typeof(IPnPSchemaSerializer).FullName) != null
                       && t.BaseType.Name == typeof(Xml.PnPBaseSchemaSerializer<>).Name)
                // Filter out those that are not targeting the current schema version or that are not in scope Template
                .Where(t =>
                {
                    var a = t.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                    return (a.MinimalSupportedSchemaVersion <= currentSchemaVersion && a.Scope == scope);
                })
                // Order the remainings by supported schema version descendant, to get first the newest ones
                .OrderByDescending(s =>
                {
                    var a = s.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                    return (a.MinimalSupportedSchemaVersion);
                }
                )
                // Group those with the same target type (which is the first generic Type argument)
                .GroupBy(t => t.BaseType.GenericTypeArguments.FirstOrDefault()?.FullName)
                // Order the result by SerializationSequence
                .OrderBy(g =>
                {
                    var maxInGroup = g.OrderByDescending(s =>
                    {
                        var a = s.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault();
                        return (a.MinimalSupportedSchemaVersion);
                    }
                    ).FirstOrDefault();
                    return sortingSelector(maxInGroup.GetCustomAttributes<TemplateSchemaSerializerAttribute>(false).FirstOrDefault());
                });
            return serializers;
        }

        private static void InvokeSerializers(ProvisioningTemplate template, object persistenceTemplate,
            IOrderedEnumerable<IGrouping<string, Type>> serializers, SerializationAction action)
        {
            foreach (var group in serializers)
            {
                // Get the first serializer only for each group (i.e. the most recent one for the current schema)
                var serializerType = group.FirstOrDefault();
                if (serializerType != null)
                {
                    // Create an instance of the serializer
                    var serializer = Activator.CreateInstance(serializerType) as IPnPSchemaSerializer;
                    if (serializer != null)
                    {
                        // And run the Deserialize/Serialize method
                        if (action == SerializationAction.Serialize)
                        {
                            serializer.Serialize(template, persistenceTemplate);
                        }
                        else
                        {
                            serializer.Deserialize(persistenceTemplate, template);
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Defines the action to execute with a pool of serializers
    /// </summary>
    internal enum SerializationAction
    {
        /// <summary>
        /// Will serialize content
        /// </summary>
        Serialize,
        /// <summary>
        /// Will deserialize content
        /// </summary>
        Deserialize
    }
}
