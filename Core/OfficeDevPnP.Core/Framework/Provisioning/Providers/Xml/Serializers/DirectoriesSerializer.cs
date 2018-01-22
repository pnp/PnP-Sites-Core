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
    /// Class to serialize/deserialize the Directories
    /// </summary>
    [TemplateSchemaSerializer(SerializationSequence = 250, DeserializationSequence = 250,
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201605,
        Default = true)]
    internal class DirectoriesSerializer : PnPBaseSchemaSerializer<Directory>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var filesCollection = persistence.GetPublicInstancePropertyValue("Files");

            if (filesCollection != null)
            {
                var directories = filesCollection.GetPublicInstancePropertyValue("Directory");

                var expressions = new Dictionary<Expression<Func<Directory, Object>>, IResolver>();
                expressions.Add(c => c.Level, new FromStringToEnumValueResolver(typeof(FileLevel)));

                expressions.Add(f => f.Security, new PropertyObjectTypeResolver<File>(fl => fl.Security, 
                    fl => fl.GetPublicInstancePropertyValue("Security")?.GetPublicInstancePropertyValue("BreakRoleInheritance")));
                expressions.Add(f => f.Security.RoleAssignments, new RoleAssigmentsFromSchemaToModelTypeResolver());

                template.Directories.AddRange(
                    PnPObjectsMapper.MapObjects<Directory>(directories,
                        new CollectionFromSchemaToModelTypeResolver(typeof(Directory)),
                        expressions,
                        recursive: true)
                        as IEnumerable<Directory>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.Directories != null && template.Directories.Count > 0)
            {
                var baseNamespace = PnPSerializationScope.Current?.BaseSchemaNamespace;
                var filesTypeName = $"{baseNamespace}.ProvisioningTemplateFiles, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var filesType = Type.GetType(filesTypeName, true);
                var directoryTypeName = $"{baseNamespace}.Directory, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var directoryType = Type.GetType(directoryTypeName, true);
                var fileLevelTypeName = $"{baseNamespace}.FileLevel, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var fileLevelType = Type.GetType(fileLevelTypeName, true);
                var objectSecurityTypeName = $"{baseNamespace}.ObjectSecurity, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var objectSecurityType = Type.GetType(objectSecurityTypeName, true);

                var expressions = new Dictionary<string, IResolver>();

                var dictionaryItemTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var dictionaryItemType = Type.GetType(dictionaryItemTypeName, true);
                var dictionaryItemKeySelector = CreateSelectorLambda(dictionaryItemType, "Key");
                var dictionaryItemValueSelector = CreateSelectorLambda(dictionaryItemType, "Value");

                expressions.Add($"{directoryType}.Level", new FromStringToEnumValueResolver(fileLevelType));
                expressions.Add($"{directoryType}.LevelSpecified", new ExpressionValueResolver(() => true));

                expressions.Add($"{directoryType}.Security", new PropertyObjectTypeResolver(objectSecurityType, "Security"));
                expressions.Add($"{objectSecurityType}.BreakRoleInheritance", new RoleAssignmentsFromModelToSchemaTypeResolver());

                var filesCollection = persistence.GetPublicInstancePropertyValue("Files");
                if (filesCollection == null)
                {
                    persistence.GetPublicInstanceProperty("Files").SetValue(persistence, Activator.CreateInstance(filesType, true));
                    filesCollection = persistence.GetPublicInstancePropertyValue("Files");
                }

                filesCollection.GetPublicInstanceProperty("Directory").SetValue(
                    filesCollection,
                    PnPObjectsMapper.MapObjects(template.Directories,
                    new CollectionFromModelToSchemaTypeResolver(directoryType), expressions, true));
            }
        }
    }
}
