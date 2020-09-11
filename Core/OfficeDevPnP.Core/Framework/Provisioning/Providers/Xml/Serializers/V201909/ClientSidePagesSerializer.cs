using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201805;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers.V201909
{
    /// <summary>
    /// Class to serialize/deserialize the client side pages
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201909,
        SerializationSequence = 2300, DeserializationSequence = 2300,
        Scope = SerializerScope.ProvisioningTemplate)]
    internal class ClientSidePagesSerializer : PnPBaseSchemaSerializer<ClientSidePage>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var clientSidePages = persistence.GetPublicInstancePropertyValue("ClientSidePages");

            if (clientSidePages != null)
            {
                var expressions = new Dictionary<Expression<Func<ClientSidePage, Object>>, IResolver>();

                // Manage CanvasControlProperties for CanvasControl and FieldValues for ClientSidePage
                var stringDictionaryTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var stringDictionaryType = Type.GetType(stringDictionaryTypeName, true);
                var stringDictionaryKeySelector = CreateSelectorLambda(stringDictionaryType, "Key");
                var stringDictionaryValueSelector = CreateSelectorLambda(stringDictionaryType, "Value");

                expressions.Add(cp => cp.Sections[0].Controls[0].ControlProperties,
                    new FromArrayToDictionaryValueResolver<String, String>(
                        stringDictionaryType, stringDictionaryKeySelector, stringDictionaryValueSelector));

                // FieldValues
                expressions.Add(cp => cp.FieldValues,
                    new FromArrayToDictionaryValueResolver<String, String>(
                        stringDictionaryType, stringDictionaryKeySelector, stringDictionaryValueSelector, "FieldValues"));

                // Properties
                expressions.Add(cp => cp.Properties,
                    new FromArrayToDictionaryValueResolver<String, String>(
                        stringDictionaryType, stringDictionaryKeySelector, stringDictionaryValueSelector, "Properties"));

                // Manage WebPartType for CanvasControl
                expressions.Add(cp => cp.Sections[0].Controls[0].Type,
                    new ExpressionValueResolver(
                        (s, p) => (Model.WebPartType)Enum.Parse(typeof(Model.WebPartType), s.GetPublicInstancePropertyValue("WebPartType").ToString())
                        ));

                // Manage ControlId for CanvasControl
                expressions.Add(cp => cp.Sections[0].Controls[0].ControlId,
                    new FromStringToGuidValueResolver());

                // Manage Header for client side page
                expressions.Add(cp => cp.Header, 
                    new Resolvers.V201909.ClientSidePageHeaderFromSchemaToModelTypeResolver());

                // Manage Security for client side page
                expressions.Add(cp => cp.Security, new PropertyObjectTypeResolver<File>(fl => fl.Security,
                    fl => fl.GetPublicInstancePropertyValue("Security")?.GetPublicInstancePropertyValue("BreakRoleInheritance")));
                expressions.Add(cp => cp.Security.RoleAssignments, new RoleAssigmentsFromSchemaToModelTypeResolver());

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
            if (template.ClientSidePages != null && template.ClientSidePages.Count > 0)
            {
                var baseClientSidePageTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.BaseClientSidePage, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var baseClientSidePageType = Type.GetType(baseClientSidePageTypeName, true);
                var clientSidePageTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ClientSidePage, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var clientSidePageType = Type.GetType(clientSidePageTypeName, true);
                var canvasSectionTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CanvasSection, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var canvasSectionType = Type.GetType(canvasSectionTypeName, true);
                var canvasControlTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CanvasControl, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var canvasControlType = Type.GetType(canvasControlTypeName, true);
                var canvasControlWebPartTypeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CanvasControlWebPartType, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var canvasControlWebPartTypeType = Type.GetType(canvasControlWebPartTypeTypeName, true);
                var objectSecurityTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ObjectSecurity, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var objectSecurityType = Type.GetType(objectSecurityTypeName, true);

                var expressions = new Dictionary<string, IResolver>();

                // Manage PromoteAsNewsArticleSpecified property for ClientSidePage
                expressions.Add($"{baseClientSidePageType}.PromoteAsNewsArticleSpecified", new ExpressionValueResolver((s, p) => true));

                // Manage PromoteAsNewsArticleSpecified property for ClientSidePage
                expressions.Add($"{baseClientSidePageType}.OverwriteSpecified", new ExpressionValueResolver((s, p) => true));

                // Manage OrderSpecified property for CanvasZone
                expressions.Add($"{canvasSectionType}.OrderSpecified", new ExpressionValueResolver((s, p) => true));

                // Manage TypeSpecified property for CanvasZone
                expressions.Add($"{canvasSectionType}.TypeSpecified", new ExpressionValueResolver((s, p) => true));

                // Manage WebPartType for CanvasControl
                expressions.Add($"{canvasControlType}.WebPartType",
                    new ExpressionValueResolver(
                        (s, p) => Enum.Parse(canvasControlWebPartTypeType, s.GetPublicInstancePropertyValue("Type").ToString()))
                        );

                // Manage CanvasControlProperties for CanvasControl and FieldValues for ClientSidePage
                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");

                expressions.Add($"{canvasControlType}.CanvasControlProperties",
                    new FromDictionaryToArrayValueResolver<string, string>(
                        dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "ControlProperties"));

                expressions.Add($"{baseClientSidePageType}.FieldValues", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                expressions.Add($"{baseClientSidePageType}.Properties", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                // Manage Header for client side page
                var clientSidePageHeaderType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.BaseClientSidePageHeader, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", false);

                if (null != clientSidePageHeaderType)
                {
                    expressions.Add($"{baseClientSidePageType}.Header", new Resolvers.V201909.ClientSidePageHeaderFromModelToSchemaTypeResolver());
                    expressions.Add($"{clientSidePageHeaderType}.TranslateX", new FromNullableToSpecifiedValueResolver<double>("TranslateXSpecified"));
                    expressions.Add($"{clientSidePageHeaderType}.TranslateY", new FromNullableToSpecifiedValueResolver<double>("TranslateYSpecified"));
                }

                // Manage Security for client side page
                expressions.Add($"{baseClientSidePageType}.Security", new Resolvers.V201807.ClientSidePageSecurityFromModelToSchemaTypeResolver());
                expressions.Add($"{objectSecurityType}.BreakRoleInheritance", new RoleAssignmentsFromModelToSchemaTypeResolver());

                // Force the specified property for LCID
                expressions.Add($"{clientSidePageType}.LCIDSpecified", new ExpressionValueResolver(((s, p) => {
                    var csp = s as ClientSidePage;
                    if (csp != null)
                    {
                        return (csp.LCID > 0);
                    }
                    else
                    {
                        return (false);
                    }
                })));

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
}
