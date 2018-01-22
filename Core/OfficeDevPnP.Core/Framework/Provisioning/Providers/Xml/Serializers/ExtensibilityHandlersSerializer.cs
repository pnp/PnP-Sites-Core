using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Linq.Expressions;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Providers for Extensibility
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 2100, DeserializationSequence = 2100,
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
                expressions.Add(h => h.Configuration, new ExpressionValueResolver((s, v) => (v as XmlElement)?.ToProviderConfiguration()));
                expressions.Add(h => h.Type, new ExpressionValueResolver<ExtensibilityHandler>((s, v, d) =>
                {
                    string res = null;
                    var typeName = s.GetPublicInstancePropertyValue("HandlerType");
                    if(typeName != null)
                    {
                        var type = Type.GetType(typeName.ToString(), false);
                        if(type != null)
                        {
                            d.Assembly = type.Assembly.FullName;
                            res = type.FullName;
                        }
                    }
                    return res;
                }));


                var result = PnPObjectsMapper.MapObjects<ExtensibilityHandler>(handlers,
                        new CollectionFromSchemaToModelTypeResolver(typeof(ExtensibilityHandler)),
                        expressions,
                        recursive: true)
                        as IEnumerable<ExtensibilityHandler>;
                template.ExtensibilityHandlers.AddRange(result.Where(h => !string.IsNullOrEmpty(h.Type)));
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ExtensibilityHandlers != null && template.ExtensibilityHandlers.Count > 0)
            {
                var providerType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Provider, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);

                var expressions = new Dictionary<string, IResolver>();
                expressions.Add($"{providerType}.HandlerType", new ExpressionValueResolver<ExtensibilityHandler>((s, v) => $"{s.Type}, {s.Assembly}"));
                expressions.Add($"{providerType}.Configuration", new ExpressionValueResolver<string>((v) => v?.ToXmlElement()));

                persistence.GetPublicInstanceProperty("Providers").SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.ExtensibilityHandlers, new CollectionFromModelToSchemaTypeResolver(providerType), expressions, true));
            }
        }
    }
}
