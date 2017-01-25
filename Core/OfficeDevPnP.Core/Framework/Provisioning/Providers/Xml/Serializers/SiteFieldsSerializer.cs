using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    [SupportedTemplateSchemas(Schemas = SupportedSchema.V201605 | SupportedSchema.V201703)]
    public class SiteFieldsSerializer : ISchemaSerializer
    {
        public XElement FromProvisioningTemplate(ProvisioningTemplate template, XNamespace ns)
        {
            if (template.SiteFields.Any())
            {
                var siteFieldsElement = new XElement(ns + "SiteFields");

                foreach (var field in template.SiteFields)
                {
                    siteFieldsElement.Add(new XElement(ns + "Field").Value = field.SchemaXml);
                }
                return siteFieldsElement;
            }
            return null;
        }

        public ProvisioningTemplate ToProvisioningTemplate(XElement templateElement, XNamespace ns, ProvisioningTemplate template)
        {
            var siteFieldsElement = templateElement.Elements(ns + "SiteFields").FirstOrDefault();
            if (siteFieldsElement != null)
            {
                var siteFields = siteFieldsElement.Elements("Field");
                if (siteFields.Any())
                {
                    foreach (var siteField in siteFields)
                    {
                        template.SiteFields.Add(new Field()
                        {
                            SchemaXml = siteField.ToString()
                        });
                    }
                }
            }
            return template;
        }
    }
}
