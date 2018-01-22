using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    /// <summary>
    /// Resolves a list of Views from Schema to Domain Model
    /// </summary>
    internal class ListInstanceDataRowsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => true;


        public ListInstanceDataRowsFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new Model.DataRowCollection(null);

            var dataRowsContainer = source.GetPublicInstancePropertyValue("DataRows");
            if (null != dataRowsContainer)
            {
                var dataRows = dataRowsContainer.GetPublicInstancePropertyValue("DataRow");

                result.UpdateBehavior = (Model.UpdateBehavior)Enum.Parse(typeof(Model.UpdateBehavior), dataRowsContainer.GetPublicInstancePropertyValue("UpdateBehavior")?.ToString());
                result.KeyColumn = (String)dataRowsContainer.GetPublicInstancePropertyValue("KeyColumn");

                if (null != dataRows)
                {
                    var expressions = new Dictionary<Expression<Func<Model.DataRow, Object>>, IResolver>();

                    // Define custom resolvers for DataRows Values and Security
                    var dataRowValueTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DataValue, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var dataRowValueType = Type.GetType(dataRowValueTypeName, true);
                    var dataRowValueKeySelector = CreateSelectorLambda(dataRowValueType, "FieldName");
                    var dataRowValueValueSelector = CreateSelectorLambda(dataRowValueType, "Value");
                    expressions.Add(dr => dr.Values,
                        new FromArrayToDictionaryValueResolver<String, String>(
                            dataRowValueType, dataRowValueKeySelector, dataRowValueValueSelector));

                    expressions.Add(dr => dr.Security, new SecurityFromSchemaToModelTypeResolver());

                    expressions.Add(dr => dr.Key, 
                        new ExpressionValueResolver(((s, p) => (String)s.GetPublicInstancePropertyValue("Key"))));

                    result.AddRange(
                        PnPObjectsMapper.MapObjects<Model.DataRow>(dataRows,
                            new CollectionFromSchemaToModelTypeResolver(typeof(Model.DataRow)),
                            expressions,
                            recursive: true)
                            as IEnumerable<Model.DataRow>);
                }
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
