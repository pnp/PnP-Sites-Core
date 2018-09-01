//using System;
//using System.Collections;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Xml.Linq;
//using System.Xml.Serialization;
//using OfficeDevPnP.Core.Framework.Provisioning.Model;
//using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers;
//using OfficeDevPnP.Core.Utilities;

//namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
//{
//    /// <summary>
//    /// Base class for serialization/deserialization of provisioning types
//    /// with the new schema serializer
//    /// </summary>
//    /// <typeparam name="TSchemaTemplate"></typeparam>
//    internal abstract class XmlPnPSchemaBaseProvisioningSerializer<TSchemaTemplate> : XmlPnPSchemaBaseSerializer<TSchemaTemplate>, IProvisioningFormatter
//        where TSchemaTemplate : new()
//    {
//        public XmlPnPSchemaBaseProvisioningSerializer(Stream referenceSchema):
//            base(referenceSchema)
//        {
//        }

//        public bool IsValidProvisioning(Stream provisioning)
//        {
//            // So far we simply need to rely on the other IsValid method
//            return (this.IsValid(provisioning));
//        }

//        public Stream ToFormattedProvisioning(Model.ProvisioningHierarchy provisioning)
//        {
//            throw new NotImplementedException();
//        }

//        public Model.ProvisioningHierarchy ToProvisioning(Stream template)
//        {
//            using (var scope = new PnPSerializationScope(typeof(TSchemaTemplate)))
//            {
//                // Prepare a variable to hold the resulting Provisioning instance
//                var result = new Model.ProvisioningHierarchy();

//                // Prepare a variable to hold the single source formatted template
//                var source = ProcessProvisioningStream(template, result);

//                DeserializeTemplate(source, result);

//                return (result);
//            }
//        }

//        protected Object ProcessProvisioningStream(Stream provisioning, Model.ProvisioningHierarchy result)
//        {
//            if (provisioning == null)
//            {
//                throw new ArgumentNullException(nameof(provisioning));
//            }

//            // Crate a copy of the source stream
//            MemoryStream sourceStream = new MemoryStream();
//            provisioning.CopyTo(sourceStream);
//            sourceStream.Position = 0;

//            // Check the provided template against the XML schema
//            if (!this.IsValidProvisioning(sourceStream))
//            {
//                // TODO: Use resource file
//                throw new ApplicationException("The provided provisioning file is not valid!");
//            }

//            sourceStream.Position = 0;
//            XDocument xml = XDocument.Load(sourceStream);
//            XNamespace pnp = this.NamespaceUri;

//            // Prepare a variable to hold the single source formatted template
//            TSchemaTemplate source = default(TSchemaTemplate);

//            // Determine if the root element is a Provisioning element
//            if (xml.Root.Name != pnp + "Provisioning")
//            {
//                // TODO: Use resource file
//                throw new ApplicationException("The content of the provided provisioning file is not valid! It is missing the Provisioning root element!");
//            }

//            // Deserialize the whole wrapper
//            Object wrapper = null;

//            var wrapperType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Provisioning, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
//            XmlSerializer xmlSerializer = new XmlSerializer(wrapperType);
//            using (var reader = xml.Root.CreateReader())
//            {
//                wrapper = xmlSerializer.Deserialize(reader);
//            }

//            // Handle the Parameters of the schema wrapper, if any
//            var tps = new TemplateParametersSerializer();
//            tps.Deserialize(wrapper, result);

//            // Handle the Localizations of the schema wrapper, if any
//            var ls = new LocalizationsSerializer();
//            ls.Deserialize(wrapper, result);

//            // Handle the Tenant-wide settings of the schema wrapper, if any
//            var ts = new TenantSerializer();
//            ts.Deserialize(wrapper, result);

//            // TODO: Add Sequence serializer

//            // Get the list of templates, if any, wrapped by the wrapper
//            var wrapperTemplates = wrapperType.GetProperty("Templates",
//                System.Reflection.BindingFlags.Instance |
//                System.Reflection.BindingFlags.Public |
//                System.Reflection.BindingFlags.IgnoreCase).GetValue(wrapper);

//            if (wrapperTemplates != null)
//            {
//                // Process every single Provisioning Template
//                foreach (var templates in (IEnumerable)wrapperTemplates)
//                {
//                    // Let's see if we have in-place provisioning templates to process
//                    var provisioningTemplates = templates.GetType()
//                        .GetProperty("ProvisioningTemplate",
//                            System.Reflection.BindingFlags.Instance |
//                            System.Reflection.BindingFlags.Public |
//                            System.Reflection.BindingFlags.IgnoreCase).GetValue(templates);

//                    if (provisioningTemplates != null)
//                    {
//                        foreach (var t in (IEnumerable)provisioningTemplates)
//                        {
//                            var templateId = (String)t.GetType().GetProperty("ID",
//                                System.Reflection.BindingFlags.Instance |
//                                System.Reflection.BindingFlags.Public |
//                                System.Reflection.BindingFlags.IgnoreCase).GetValue(t);

//                            if (templateId != null)
//                            {
//                                // Process the current template
//                                var template = this.ToProvisioningTemplate(provisioning, templateId, true);
//                                template.ParentProvisioningHierarchy = result;

//                                // And add it to the collection of templates for the current Provisioning object
//                                result.Templates.Add(template);
//                            }
//                        }
//                    }

//                    var provisioningTemplateFiles = templates.GetType()
//                        .GetProperty("ProvisioningTemplateFile",
//                            System.Reflection.BindingFlags.Instance |
//                            System.Reflection.BindingFlags.Public |
//                            System.Reflection.BindingFlags.IgnoreCase).GetValue(templates);

//                    // If there are external file references
//                    if (provisioningTemplateFiles != null)
//                    {
//                        foreach (var f in (IEnumerable)provisioningTemplateFiles)
//                        {
//                            var templateId = (String)f.GetType().GetProperty("ID",
//                                System.Reflection.BindingFlags.Instance |
//                                System.Reflection.BindingFlags.Public |
//                                System.Reflection.BindingFlags.IgnoreCase).GetValue(f);

//                            if (templateId != null)
//                            {
//                                // Let's see if we have an external file for the template
//                                var externalFile = (String)f.GetType().GetProperty("File",
//                                    System.Reflection.BindingFlags.Instance |
//                                    System.Reflection.BindingFlags.Public |
//                                    System.Reflection.BindingFlags.IgnoreCase).GetValue(f);

//                                Stream externalFileStream = this.Provider.Connector.GetFileStream(externalFile);

//                                // Process the template stream
//                                var template = this.ToProvisioningTemplate(externalFileStream, templateId, true);
//                                template.ParentProvisioningHierarchy = result;

//                                // And add it to the collection of templates for the current Provisioning object
//                                result.Templates.Add(template);
//                            }
//                        }
//                    }
//                }
//            }

//            return (source);
//        }
//    }
//}
