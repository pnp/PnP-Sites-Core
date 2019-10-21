using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class FoldersFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new List<Model.Folder>();

            var folders = source.GetPublicInstancePropertyValue("Folders");
            if (null == folders)
            {
                folders = source.GetPublicInstancePropertyValue("Folder1");
            }

            resolvers = new Dictionary<string, IResolver>();
            resolvers.Add($"{typeof(Model.Folder).FullName}.Folders", new FoldersFromSchemaToModelTypeResolver());
            resolvers.Add($"{typeof(Model.Folder).FullName}.Security", new SecurityFromSchemaToModelTypeResolver());

            // DefaultColumnValues
            var defaultColumnValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var defaultColumnValueType = Type.GetType(defaultColumnValueTypeName, true);
            var defaultColumnValueKeySelector = CreateSelectorLambda(defaultColumnValueType, "Key");
            var defaultColumnValueValueSelector = CreateSelectorLambda(defaultColumnValueType, "Value");
            resolvers.Add($"{typeof(Model.Folder).FullName}.DefaultColumnValues", new FromArrayToDictionaryValueResolver<string, string>(defaultColumnValueType, defaultColumnValueKeySelector, defaultColumnValueValueSelector));

            // Folders' Properties
            var folderPropertyTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.StringDictionaryItem, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var folderPropertyType = Type.GetType(folderPropertyTypeName, true);
            var folderPropertyKeySelector = CreateSelectorLambda(folderPropertyType, "Key");
            var folderPropertyValueSelector = CreateSelectorLambda(folderPropertyType, "Value");
            resolvers.Add($"{typeof(Model.Folder).FullName}.Properties",
                new FromArrayToDictionaryValueResolver<string, string>(
                    folderPropertyType, folderPropertyKeySelector, folderPropertyValueSelector));

            if (null != folders)
            {
                foreach (var f in ((IEnumerable)folders))
                {
                    var targetItem = new Model.Folder();
                    PnPObjectsMapper.MapProperties(f, targetItem, resolvers, recursive);
                    result.Add(targetItem);
                }
            }

            return (result);
        }

        private LambdaExpression CreateSelectorLambda(Type targetType, String propertyName)
        {
            return (Expression.Lambda(
                Expression.Convert(
                    Expression.MakeMemberAccess(
                        Expression.Parameter(targetType, "i"),
                        targetType.GetProperty(propertyName,
                            System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.Public)),
                    typeof(object)),
                ParameterExpression.Parameter(targetType, "i")));
        }

    }
}
