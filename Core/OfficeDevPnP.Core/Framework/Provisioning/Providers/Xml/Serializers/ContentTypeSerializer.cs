//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Xml.Linq;
//using OfficeDevPnP.Core.Framework.Provisioning.Model;

//namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
//{
//    [SupportedTemplateSchemas(Schemas = SupportedSchema.V201605 | SupportedSchema.V201703)]

//    class ContentTypeSerializer : ISchemaSerializer
//    {
//        public XElement FromProvisioningTemplate(ProvisioningTemplate template, XNamespace ns)
//        {

//        }

//        public ProvisioningTemplate ToProvisioningTemplate(XElement templateElement, XNamespace ns, ProvisioningTemplate template)
//        {
//            var contentTypesElement = templateElement.Elements(ns + "ContentTypes").FirstOrDefault();
//            if (contentTypesElement != null)
//            {
//                var contenttypes = contentTypesElement.Elements(ns + "ContentType");
//                if (contenttypes.Any())
//                {
//                    foreach (var contenttype in contenttypes)
//                    {
//                        var ct = new ContentType()
//                        {
//                            Id = contenttype.Attribute(ns + nameof(ContentType.Id).ToUpper()).Value, // Required
//                            Name = contenttype.Attribute(ns + nameof(ContentType.Name)).Value, // Required
//                            Description = contenttype.Attribute(ns + nameof(ContentType.Description))?.Value,
//                            Group = contenttype.Attribute(ns + nameof(ContentType.Group))?.Value,
//                            Hidden = contenttype.GetOptionalBoolValue(ns + nameof(ContentType.Hidden), false),
//                            Sealed = contenttype.GetOptionalBoolValue(ns + nameof(ContentType.Sealed), false),
//                            ReadOnly = contenttype.GetOptionalBoolValue(ns + nameof(ContentType.ReadOnly), false),
//                            Overwrite = contenttype.GetOptionalBoolValue(ns + nameof(ContentType.Overwrite), false),
//                            NewFormUrl = contenttype.Attribute(ns + nameof(ContentType.NewFormUrl))?.Value,
//                            EditFormUrl = contenttype.Attribute(ns + nameof(ContentType.EditFormUrl))?.Value,
//                            DisplayFormUrl = contenttype.Attribute(ns + nameof(ContentType.DisplayFormUrl))?.Value,
//                        };
//                        var fieldRefElements = contenttype.Descendants(ns + "FieldRef");
//                        if (fieldRefElements.Any())
//                        {
//                            foreach (var fieldRefElement in fieldRefElements)
//                            {
//                                var name = fieldRefElement.Attribute(ns + nameof(FieldRef.Name)).Value;
//                                var fieldRef = new FieldRef(name);
//                                fieldRef.DisplayName = fieldRefElement.Attribute(ns + nameof(FieldRef.DisplayName))?.Value;
//                                fieldRef.Hidden = fieldRefElement.GetOptionalBoolValue(ns + nameof(FieldRef.Hidden),false);
//                                fieldRef.Required = fieldRefElement.GetOptionalBoolValue(ns + nameof(FieldRef.Required), false);
//                                fieldRef.Id 

//                            }
//                        }

//                        template.ContentTypes.Add(ct);
//                    }
//                }
//            }
//            return template;
//        }
//    }
