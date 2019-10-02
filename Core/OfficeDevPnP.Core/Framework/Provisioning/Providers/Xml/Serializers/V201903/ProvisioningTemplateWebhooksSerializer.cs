using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Provisioning Webhooks
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201903,
        SerializationSequence = 2500, DeserializationSequence = 2500,
        Scope = SerializerScope.ProvisioningTemplate)]
    internal class ProvisioningTemplateWebhooksSerializer : PnPBaseSchemaSerializer<ProvisioningTemplateWebhook>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var provisioningTemplateWebhooks = persistence.GetPublicInstancePropertyValue("ProvisioningTemplateWebhooks");

            if (provisioningTemplateWebhooks != null)
            {
                var expressions = new Dictionary<Expression<Func<ProvisioningTemplateWebhook, Object>>, IResolver>();
                
                // Parameters
                var parameterTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var parameterType = Type.GetType(parameterTypeName, true);
                var parameterKeySelector = CreateSelectorLambda(parameterType, "Key");
                var parameterValueSelector = CreateSelectorLambda(parameterType, "Value");
                expressions.Add(w => w.Parameters,
                    new FromArrayToDictionaryValueResolver<string, string>(
                        parameterType, parameterKeySelector, parameterValueSelector));

                template.ProvisioningTemplateWebhooks.AddRange(
                    PnPObjectsMapper.MapObjects(provisioningTemplateWebhooks,
                            new CollectionFromSchemaToModelTypeResolver(typeof(ProvisioningTemplateWebhook)),
                            expressions, recursive: true)
                            as IEnumerable<ProvisioningTemplateWebhook>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ProvisioningTemplateWebhooks != null && template.ProvisioningTemplateWebhooks.Count > 0)
            {
                var provisioningTemplateWebhookTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ProvisioningTemplateWebhooksProvisioningTemplateWebhook, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var provisioningTemplateWebhookType = Type.GetType(provisioningTemplateWebhookTypeName, true);

                var expressions = new Dictionary<string, IResolver>();

                // Parameters
                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");

                expressions.Add($"{provisioningTemplateWebhookType}.Parameters", 
                    new FromDictionaryToArrayValueResolver<string, string>(
                        dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                persistence.GetPublicInstanceProperty("ProvisioningTemplateWebhooks")
                    .SetValue(
                        persistence,
                        PnPObjectsMapper.MapObjects(template.ProvisioningTemplateWebhooks,
                            new CollectionFromModelToSchemaTypeResolver(provisioningTemplateWebhookType),
                            expressions,
                            recursive: true));
            }
        }
    }
}
