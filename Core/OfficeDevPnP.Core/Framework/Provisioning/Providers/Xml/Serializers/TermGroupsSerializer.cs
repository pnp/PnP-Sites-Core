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
    internal class TermGroupsSerializer : PnPBaseSchemaSerializer<TermGroup>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var groups = persistence.GetPublicInstancePropertyValue("TermGroups");

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
                new PropertyObjectTypeResolver<Term>(t=>t.Terms, 
                v => v.GetPublicInstancePropertyValue("Terms")?.GetPublicInstancePropertyValue("Items"),
                new CollectionFromSchemaToModelTypeResolver(typeof(Term))));

            //expressions.Add(p => p.Layout, new FromStringToEnumValueResolver(typeof(WikiPageLayout)));

            template.TermGroups.AddRange(
                PnPObjectsMapper.MapObjects<TermGroup>(groups,
                    new CollectionFromSchemaToModelTypeResolver(typeof(TermGroup)),
                    expressions,
                    recursive: true)
                    as IEnumerable<TermGroup>);
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Pages != null && template.Pages.Count > 0)
            {
                var baseNamespace = PnPSerializationScope.Current?.BaseSchemaNamespace;
                var pageTypeName = $"{baseNamespace}.Page, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var pageType = Type.GetType(pageTypeName, true);
                var layoutTypeName = $"{baseNamespace}.WikiPageLayout, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var layoutType = Type.GetType(layoutTypeName, true);
                var objectSecurityTypeName = $"{baseNamespace}.ObjectSecurity, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var objectSecurityType = Type.GetType(objectSecurityTypeName, true);

                var expressions = new Dictionary<string, IResolver>();

                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.BaseFieldValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "FieldName");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");

                expressions.Add($"{pageType}.Fields", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));
                expressions.Add($"{pageType}.Layout", new FromStringToEnumValueResolver(layoutType));
                expressions.Add($"{pageType}.LayoutSpecified", new ExpressionValueResolver((s, v) => true));

                expressions.Add($"{pageType}.Security", new PropertyObjectTypeResolver(objectSecurityType, "Security"));
                expressions.Add($"{objectSecurityType}.BreakRoleInheritance", new RoleAssigmentsFromModelToSchemaTypeResolver());

                expressions.Add($"{baseNamespace}.WikiPageWebPart.Row", new ExpressionValueResolver((s, v) => (int)(uint)v));
                expressions.Add($"{baseNamespace}.WikiPageWebPart.Column", new ExpressionValueResolver((s, v) => (int)(uint)v));
                //convert webpart content to xml element
                expressions.Add($"{baseNamespace}.WikiPageWebPart.Contents", new ExpressionValueResolver((s, v) =>
                {
                    var doc = new XmlDocument();
                    var str = v != null ? v.ToString() : null;
                    if (!string.IsNullOrEmpty(str))
                    {
                        doc.LoadXml(str);
                    }
                    return doc.DocumentElement;
                }));

                //persistence.GetPublicInstanceProperty("Pages").SetValue(
                    //persistence,
                    //PnPObjectsMapper.MapObjects(template.Pages,
                    //new CollectionFromModelToSchemaTypeResolver(pageType), expressions, true));
            }
        }
    }
}
