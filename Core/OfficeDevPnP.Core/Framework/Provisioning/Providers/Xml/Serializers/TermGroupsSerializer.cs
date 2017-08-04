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
    /// Class to serialize/deserialize the Term Groups
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 1600, DeserializationSequence = 1600,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class TermGroupsSerializer : PnPBaseSchemaSerializer<TermGroup>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var groups = persistence.GetPublicInstancePropertyValue("TermGroups");

            if (groups != null)
            {
                var expressions = new Dictionary<Expression<Func<TermGroup, Object>>, IResolver>();
                expressions.Add(g => g.Id, new FromStringToGuidValueResolver());
                expressions.Add(g => g.TermSets[0].Id, new FromStringToGuidValueResolver());
                expressions.Add(g => g.TermSets[0].Terms[0].Id, new FromStringToGuidValueResolver());
                expressions.Add(g => g.TermSets[0].Terms[0].SourceTermId, new FromStringToGuidValueResolver());

                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");
                expressions.Add(g => g.TermSets[0].Properties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));
                expressions.Add(g => g.TermSets[0].Terms[0].LocalProperties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "LocalCustomProperties"));
                expressions.Add(g => g.TermSets[0].Terms[0].Properties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "CustomProperties"));
                expressions.Add(g => g.TermSets[0].Terms[0].Terms,
                    new PropertyObjectTypeResolver<Term>(t => t.Terms,
                    v => v.GetPublicInstancePropertyValue("Terms")?.GetPublicInstancePropertyValue("Items"),
                    new CollectionFromSchemaToModelTypeResolver(typeof(Term))));

                template.TermGroups.AddRange(
                    PnPObjectsMapper.MapObjects<TermGroup>(groups,
                        new CollectionFromSchemaToModelTypeResolver(typeof(TermGroup)),
                        expressions,
                        recursive: true)
                        as IEnumerable<TermGroup>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.TermGroups != null && template.TermGroups.Count > 0)
            {
                var baseNamespace = PnPSerializationScope.Current?.BaseSchemaNamespace;
                var termGroupType = Type.GetType($"{baseNamespace}.TermGroup, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var termSetType = Type.GetType($"{baseNamespace}.TermSet, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var termType = Type.GetType($"{baseNamespace}.Term, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                var termTermsType = Type.GetType($"{baseNamespace}.TermTerms, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);

                var expressions = new Dictionary<string, IResolver>();

                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");

                expressions.Add($"{termGroupType}.SiteCollectionTermGroupSpecified", new ExpressionValueResolver((s, v) => (bool)s.GetPublicInstancePropertyValue("SiteCollectionTermGroup")));
                expressions.Add($"{termGroupType}.TermSets", new TermSetFromModelToSchemaTypeResolver());

                expressions.Add($"{termSetType}.Language", new FromNullableToSpecifiedValueResolver<int>("LanguageSpecified"));
                expressions.Add($"{termSetType}.CustomProperties", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "Properties"));
                expressions.Add($"{termType}.Language", new FromNullableToSpecifiedValueResolver<int>("LanguageSpecified"));
                expressions.Add($"{termType}.CustomProperties", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "Properties"));
                expressions.Add($"{termType}.LocalCustomProperties", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector, "LocalProperties"));
                expressions.Add($"{termType}.SourceTermId", new ExpressionValueResolver<Guid>((v) => v != Guid.Empty ? v.ToString() : null));
                expressions.Add($"{termType}.Terms", new ExpressionTypeResolver<Term>(termTermsType, (source, resolvers, recursive, dest) =>
                {
                    dest.SetPublicInstancePropertyValue("Items", (new CollectionFromModelToSchemaTypeResolver(termType)).Resolve(source.Terms, resolvers, recursive));
                }));

                persistence.GetPublicInstanceProperty("TermGroups").SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.TermGroups,
                    new CollectionFromModelToSchemaTypeResolver(termGroupType), expressions, true));
            }
        }
    }
}
