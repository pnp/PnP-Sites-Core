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
    /// Class to serialize/deserialize the Files
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 1400, DeserializationSequence = 1400,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class FilesSerializer : PnPBaseSchemaSerializer<File>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var filesCollection = persistence.GetPublicInstancePropertyValue("Files");

            if (filesCollection != null)
            {
                var files = filesCollection.GetPublicInstancePropertyValue("File");

                var expressions = new Dictionary<Expression<Func<File, Object>>, IResolver>();
                expressions.Add(c => c.Level, new FromStringToEnumValueResolver(typeof(FileLevel)));

                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");
                expressions.Add(f => f.Properties, new FromArrayToDictionaryValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));
                expressions.Add(f => f.Security, new PropertyObjectTypeResolver<File>(fl => fl.Security, 
                    fl => fl.GetPublicInstancePropertyValue("Security")?.GetPublicInstancePropertyValue("BreakRoleInheritance")));
                expressions.Add(f => f.Security.RoleAssignments, new RoleAssigmentsFromSchemaToModelTypeResolver());
                expressions.Add(f => f.WebParts[0].Order, new ExpressionValueResolver<int>((v) => (uint)v));
                expressions.Add(f => f.WebParts[0].Contents, new ExpressionValueResolver((s, v) => v != null ? ((XmlElement)v).OuterXml : null));

                template.Files.AddRange(
                    PnPObjectsMapper.MapObjects<File>(files,
                        new CollectionFromSchemaToModelTypeResolver(typeof(File)),
                        expressions,
                        recursive: true)
                        as IEnumerable<File>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Files != null && template.Files.Count > 0)
            {
                var baseNamespace = PnPSerializationScope.Current?.BaseSchemaNamespace;
                var filesTypeName = $"{baseNamespace}.ProvisioningTemplateFiles, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var filesType = Type.GetType(filesTypeName, true);
                var fileTypeName = $"{baseNamespace}.File, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var fileType = Type.GetType(fileTypeName, true);
                var fileLevelTypeName = $"{baseNamespace}.FileLevel, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var fileLevelType = Type.GetType(fileLevelTypeName, true);
                var objectSecurityTypeName = $"{baseNamespace}.ObjectSecurity, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var objectSecurityType = Type.GetType(objectSecurityTypeName, true);

                var expressions = new Dictionary<string, IResolver>();

                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");

                expressions.Add($"{fileType}.Properties", new FromDictionaryToArrayValueResolver<string, string>(dictionaryItemType, dictionaryItemKeySelector, dictionaryItemValueSelector));
                expressions.Add($"{fileType}.Level", new FromStringToEnumValueResolver(fileLevelType));
                expressions.Add($"{fileType}.LevelSpecified", new ExpressionValueResolver(() => true));

                expressions.Add($"{fileType}.Security", new PropertyObjectTypeResolver(objectSecurityType, "Security"));
                expressions.Add($"{objectSecurityType}.BreakRoleInheritance", new RoleAssignmentsFromModelToSchemaTypeResolver());

                expressions.Add($"{baseNamespace}.WebPartPageWebPart.Order", new ExpressionValueResolver<uint>((v) => (int)v));
                //convert webpart content to xml element
                expressions.Add($"{baseNamespace}.WebPartPageWebPart.Contents", new ExpressionValueResolver<string>((v) => v?.ToXmlElement()));

                var filesCollection = persistence.GetPublicInstancePropertyValue("Files");
                if (filesCollection == null)
                {
                    persistence.GetPublicInstanceProperty("Files").SetValue(persistence, Activator.CreateInstance(filesType, true));
                    filesCollection = persistence.GetPublicInstancePropertyValue("Files");
                }

                filesCollection.GetPublicInstanceProperty("File").SetValue(
                    filesCollection,
                    PnPObjectsMapper.MapObjects(template.Files,
                    new CollectionFromModelToSchemaTypeResolver(fileType), expressions, true));
            }
        }
    }
}
