using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Pages
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 1500, DeserializationSequence = 1500,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class PagesSerializer : PnPBaseSchemaSerializer<Page>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var pages = persistence.GetPublicInstancePropertyValue("Pages");

            if (pages != null)
            {
                var expressions = new Dictionary<Expression<Func<Page, Object>>, IResolver>();
                expressions.Add(p => p.Layout, new FromStringToEnumValueResolver(typeof(WikiPageLayout)));

                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.BaseFieldValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "FieldName");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");
                expressions.Add(f => f.Fields, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));

                expressions.Add(f => f.Security, new PropertyObjectTypeResolver<File>(fl => fl.Security,
                    fl => fl.GetPublicInstancePropertyValue("Security")?.GetPublicInstancePropertyValue("BreakRoleInheritance")));
                expressions.Add(f => f.Security.RoleAssignments, new RoleAssigmentsFromSchemaToModelTypeResolver());
                expressions.Add(f => f.WebParts[0].Row, new ExpressionValueResolver((s, v) => (uint)(int)v));
                expressions.Add(f => f.WebParts[0].Column, new ExpressionValueResolver((s, v) => (uint)(int)v));
                //Contents is deserialized to persistence differently base on schema version. 
                //Deserialization of older schemas include <Contents> elements, newer do not. 
                //Deserialized model should not contain <Contents> element, so skip it if present
                expressions.Add(f => f.WebParts[0].Contents, new ExpressionValueResolver((s, v) => {
                    if (v != null)
                    {
                        var xml = (XmlElement)v;
                        return xml.Name == "Contents" ? xml.InnerXml : xml.OuterXml;
                    }
                    return null;
                }));

                template.Pages.AddRange(
                    PnPObjectsMapper.MapObjects<Page>(pages,
                        new CollectionFromSchemaToModelTypeResolver(typeof(Page)),
                        expressions,
                        recursive: true)
                        as IEnumerable<Page>);
            }
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
                expressions.Add($"{objectSecurityType}.BreakRoleInheritance", new RoleAssignmentsFromModelToSchemaTypeResolver());

                expressions.Add($"{baseNamespace}.WikiPageWebPart.Row", new ExpressionValueResolver<uint>((v) => (int)v));
                expressions.Add($"{baseNamespace}.WikiPageWebPart.Column", new ExpressionValueResolver<uint>((v) => (int)v));
                expressions.Add($"{baseNamespace}.WikiPageWebPart.Contents", new ExpressionValueResolver<string>((v) => v?.ToXmlElement()));

                persistence.GetPublicInstanceProperty("Pages").SetValue(
                    persistence,
                    PnPObjectsMapper.MapObjects(template.Pages,
                    new CollectionFromModelToSchemaTypeResolver(pageType), expressions, true));
            }
        }
    }
}
