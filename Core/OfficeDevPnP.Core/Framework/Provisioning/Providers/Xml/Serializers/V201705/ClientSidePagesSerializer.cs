using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Site Webhooks
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201705,
        SerializationSequence = 2300, DeserializationSequence = 2300,
        Default = true)]
    internal class ClientSidePagesSerializer : PnPBaseSchemaSerializer<ClientSidePage>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var clientSidePages = persistence.GetPublicInstancePropertyValue("ClientSidePages");

            if (clientSidePages != null)
            {
                var expressions = new Dictionary<Expression<Func<ClientSidePage, Object>>, IResolver>();

                // Manage CanvasControlProperties for CanvasControl
                var stringDictionaryTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var stringDictionaryType = Type.GetType(stringDictionaryTypeName, true);
                var stringDictionaryKeySelector = CreateSelectorLambda(stringDictionaryType, "Key");
                var stringDictionaryValueSelector = CreateSelectorLambda(stringDictionaryType, "Value");
                expressions.Add(cp => cp.Zones[0].Controls[0].ControlProperties,
                    new FromArrayToDictionaryValueResolver<String, String>(
                        stringDictionaryType, stringDictionaryKeySelector, stringDictionaryValueSelector));

                // Manage WebPartType for CanvasControl
                expressions.Add(cp => cp.Zones[0].Controls[0].Type,
                    new ExpressionValueResolver(
                        (s, p) => (Model.WebPartType)Enum.Parse(typeof(Model.WebPartType), s.GetPublicInstancePropertyValue("WebPartType").ToString())
                        ));

                // Manage ControlId for CanvasControl
                expressions.Add(cp => cp.Zones[0].Controls[0].ControlId,
                    new FromStringToGuidValueResolver());

                template.ClientSidePages.AddRange(
                    PnPObjectsMapper.MapObjects(clientSidePages,
                            new CollectionFromSchemaToModelTypeResolver(typeof(ClientSidePage)),
                            expressions,
                            recursive: true)
                        as IEnumerable<ClientSidePage>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            var clientSidePageTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ClientSidePage, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var clientSidePageType = Type.GetType(clientSidePageTypeName, true);
            var canvasZoneTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CanvasZone, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var canvasZoneType = Type.GetType(canvasZoneTypeName, true);
            var canvasControlTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CanvasControl, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var canvasControlType = Type.GetType(canvasControlTypeName, true);
            var canvasControlWebPartTypeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CanvasControlWebPartType, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var canvasControlWebPartTypeType = Type.GetType(canvasControlWebPartTypeTypeName, true);

            var expressions = new Dictionary<string, IResolver>();

            // Manage PromoteAsNewsArticleSpecified property for ClientSidePage
            expressions.Add($"{clientSidePageType}.PromoteAsNewsArticleSpecified", new ExpressionValueResolver((s, p) => true));
            
            // Manage OrderSpecified property for CanvasZone
            expressions.Add($"{canvasZoneType}.OrderSpecified", new ExpressionValueResolver((s, p) => true));

            // Manage TypeSpecified property for CanvasZone
            expressions.Add($"{canvasZoneType}.TypeSpecified", new ExpressionValueResolver((s, p) => true));

            // Manage WebPartType for CanvasControl
            expressions.Add($"{canvasControlType}.WebPartType",
                new ExpressionValueResolver(
                    (s, p) => Enum.Parse(canvasControlWebPartTypeType, s.GetPublicInstancePropertyValue("Type").ToString()))
                    );

            // Manage CanvasControlProperties for CanvasControl
            var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
            var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
            var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");

            expressions.Add($"{canvasControlType}.CanvasControlProperties", 
                new FromDictionaryToArrayValueResolver<string, string>(
                    dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "ControlProperties"));

            persistence.GetPublicInstanceProperty("ClientSidePages")
                .SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.ClientSidePages,
                        new CollectionFromModelToSchemaTypeResolver(clientSidePageType),
                        expressions,
                        recursive: true));
        }
    }
}
