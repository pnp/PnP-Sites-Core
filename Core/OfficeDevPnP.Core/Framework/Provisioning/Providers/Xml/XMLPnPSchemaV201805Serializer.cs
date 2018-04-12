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
    /// <summary>
    /// Implements the logic to serialize a schema of version 201801
    /// </summary>
    internal class XMLPnPSchemaV201805Serializer : XmlPnPSchemaBaseSerializer<V201805.ProvisioningTemplate>
    {
        public XMLPnPSchemaV201805Serializer():
            base(typeof(XMLConstants)
                .Assembly
                .GetManifestResourceStream("OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2018-05.xsd"))
        {
        }

        public override string NamespaceUri
        {
            get { return (XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2018_05); }
        }

        public override string NamespacePrefix
        {
            get { return (XMLConstants.PROVISIONING_SCHEMA_PREFIX); }
        }

        protected override void DeserializeTemplate(object persistenceTemplate, Model.ProvisioningTemplate template)
        {
            base.DeserializeTemplate(persistenceTemplate, template);
        }

        protected override void SerializeTemplate(Model.ProvisioningTemplate template, object persistenceTemplate)
        {
            base.SerializeTemplate(template, persistenceTemplate);
        }
    }
}

