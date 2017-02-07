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

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    internal abstract class XmlPnPSchemaBaseSerializer : IXMLSchemaFormatter, ITemplateFormatter
    {
        private TemplateProviderBase _provider;
        private Stream _referenceSchema;

        public XmlPnPSchemaBaseSerializer(Stream referenceSchema)
        {
            if (referenceSchema == null)
            {
                throw new ArgumentNullException("referenceSchema");
            }

            this._referenceSchema = referenceSchema;
        }

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

        public abstract string NamespacePrefix { get; }
        public abstract string NamespaceUri { get; }
        public abstract Stream ToFormattedTemplate(ProvisioningTemplate template);
        public abstract ProvisioningTemplate ToProvisioningTemplate(Stream template);
        public abstract ProvisioningTemplate ToProvisioningTemplate(Stream template, string identifier);
    }
}
