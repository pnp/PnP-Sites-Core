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
    [TemplateSchemaSerializer(SerializationSequence = 1200, DeserializationSequence = 1200,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class FeaturesSerializer : PnPBaseSchemaSerializer<Features>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var features = persistence.GetPublicInstancePropertyValue("Features");

            var expressions = new Dictionary<Expression<Func<Features, Object>>, IResolver>();
            expressions.Add(f => f.SiteFeatures[0].Id, new FromStringToGuidValueResolver());

            PnPObjectsMapper.MapProperties(features, template.Features, expressions, true);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var featuresTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Features, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var featuresType = Type.GetType(featuresTypeName, true);
            var target = Activator.CreateInstance(featuresType, true);

            PnPObjectsMapper.MapProperties(template.Features, target, null, recursive: true);

            persistence.GetPublicInstanceProperty("Features").SetValue(persistence, target);
        }
    }
}
