using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the content types
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 250, DeserializationSequence = 250,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class ExtensibilityHandlersSerializer : PnPBaseSchemaSerializer<ExtensibilityHandler>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var handlers = persistence.GetPublicInstancePropertyValue("Providers");
            if (handlers != null)
            {
                var expressions = new Dictionary<Expression<Func<ExtensibilityHandler, Object>>, IResolver>();

                template.ExtensibilityHandlers.AddRange(
                    PnPObjectsMapper.MapObjects<ExtensibilityHandler>(handlers,
                        new CollectionFromSchemaToModelTypeResolver(typeof(ExtensibilityHandler)),
                        expressions,
                        recursive: true)
                        as IEnumerable<ExtensibilityHandler>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ExtensibilityHandlers != null && template.ExtensibilityHandlers.Count > 0)
            {
                var providerType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Provider, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);

                var expressions = new Dictionary<string, IResolver>();

                persistence.GetPublicInstanceProperty("Providers").SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.ExtensibilityHandlers, new CollectionFromModelToSchemaTypeResolver(providerType), expressions, true));
            }
        }
    }
}
