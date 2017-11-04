using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    internal class ListInstanceDataRowsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public ListInstanceDataRowsFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var listInstanceDataRowsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstanceDataRows, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var listInstanceDataRowsType = Type.GetType(listInstanceDataRowsTypeName, true);

            var listInstanceDataRowTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstanceDataRowsDataRow, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var listInstanceDataRowType = Type.GetType(listInstanceDataRowTypeName, true);

            var listInstanceDataRowValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DataValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var listInstanceDataRowValueType = Type.GetType(listInstanceDataRowValueTypeName, true);

            var dataRowValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DataValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var dataRowValueType = Type.GetType(dataRowValueTypeName, true);
            var dataRowValueKeySelector = CreateSelectorLambda(dataRowValueType, "FieldName");
            var dataRowValueValueSelector = CreateSelectorLambda(dataRowValueType, "Value");

            resolvers.Add($"{listInstanceDataRowType}.DataValue", new FromDictionaryToArrayValueResolver<string, string>(dataRowValueType, dataRowValueKeySelector, dataRowValueValueSelector, "Values"));
            resolvers.Add($"{listInstanceDataRowType}.Security", new SecurityFromModelToSchemaTypeResolver());

            Object result = null;

            var list = source as Model.ListInstance;

            if (null != list && null != list.DataRows && list.DataRows.Count > 0)
            {
                // Prepare the DataRows element wrapper
                result = Activator.CreateInstance(listInstanceDataRowsType);
                
                result.GetPublicInstanceProperty("KeyColumn").SetValue(result, list.DataRows.KeyColumn);
                result.GetPublicInstanceProperty("UpdateBehavior").SetValue(result, list.DataRows.UpdateBehavior);

                // Process each DataRow element
                var dataRows = Array.CreateInstance(listInstanceDataRowType, list.DataRows.Count);

                Int32 i = 0;
                foreach (var dr in list.DataRows)
                {
                    Object dataRow = Activator.CreateInstance(listInstanceDataRowType);
                    PnPObjectsMapper.MapProperties(dr, dataRow, resolvers, recursive);
                    dataRows.SetValue(dataRow, i);
                    i++;
                }

                result.SetPublicInstancePropertyValue("DataRow", dataRows);
            }

            return (result);
        }

        /// <summary>
        /// Protected method to create a Lambda Expression like: i => i.Property
        /// </summary>
        /// <param name="targetType">The Type of the .NET property to apply the Lambda Expression to</param>
        /// <param name="propertyName">The name of the property of the target object</param>
        /// <returns></returns>
        protected LambdaExpression CreateSelectorLambda(Type targetType, String propertyName)
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
