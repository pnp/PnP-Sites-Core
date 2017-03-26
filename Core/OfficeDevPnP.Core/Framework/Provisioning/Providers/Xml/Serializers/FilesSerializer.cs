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
                expressions.Add(f => f.WebParts[0].Order, new ExpressionValueResolver((s, v) => (uint)(int)v));
                expressions.Add(f => f.WebParts[0].Contents, new ExpressionValueResolver((s, v) => v != null ? ((XmlElement)v).InnerXml : null));

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
            var baseNamespace = PnPSerializationScope.Current?.BaseSchemaNamespace;
            var filesTypeName = $"{baseNamespace}.ProvisioningTemplateFiles, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var filesType = Type.GetType(filesTypeName, true);
            var fileTypeName = $"{baseNamespace}.File, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var fileType = Type.GetType(fileTypeName, true);

            var expressions = new Dictionary<string, IResolver>();

            var filesCollection = persistence.GetPublicInstancePropertyValue("Files");
            if(filesCollection == null)
            {
                persistence.GetPublicInstanceProperty("Files").SetValue(persistence, Activator.CreateInstance(filesType, true));
                filesCollection = persistence.GetPublicInstancePropertyValue("Files");
            }

            //filesCollection.GetPublicInstanceProperty("File").SetValue(
            //    filesCollection,
            //    PnPObjectsMapper.MapObjects(template.Files,
            //    new CollectionFromModelToSchemaTypeResolver(fileType), expressions, true));

        }
    }
}
