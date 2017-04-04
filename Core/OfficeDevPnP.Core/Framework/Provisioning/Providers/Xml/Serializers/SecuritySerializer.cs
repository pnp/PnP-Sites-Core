using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Features
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 400, DeserializationSequence = 400,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class SecuritySerializer : PnPBaseSchemaSerializer<SiteSecurity>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var security = persistence.GetPublicInstancePropertyValue("Security");

            PnPObjectsMapper.MapProperties(security, template.Security, null, true);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var securityTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Security, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var securityType = Type.GetType(securityTypeName, true);
            var target = Activator.CreateInstance(securityType, true);

            PnPObjectsMapper.MapProperties(template.Security, target, null, recursive: true);

            persistence.GetPublicInstanceProperty("Security").SetValue(persistence, target);
        }
    }
}
