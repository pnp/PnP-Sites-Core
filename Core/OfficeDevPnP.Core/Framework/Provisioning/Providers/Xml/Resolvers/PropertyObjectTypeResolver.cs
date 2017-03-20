using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Typed vesion of PropertyObjectTypeResolver
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class PropertyObjectTypeResolver<T> : PropertyObjectTypeResolver
    {
        public PropertyObjectTypeResolver(Expression<Func<T, object>> exp) : base(GetPropertyType(exp), GetPropertyName(exp))
        {
        }

        private static string GetPropertyName(Expression<Func<T, object>> exp)
        {
            return (exp.Body as MemberExpression ?? ((UnaryExpression)exp.Body).Operand as MemberExpression).Member.Name;
        }

        private static Type GetPropertyType(Expression<Func<T, object>> exp)
        {
            return (exp.Body as MemberExpression ?? ((UnaryExpression)exp.Body).Operand as MemberExpression).Type;
        }
    }

    /// <summary>
    /// Resolves a collection type from Domain Model to Schema
    /// </summary>
    internal class PropertyObjectTypeResolver : ITypeResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        private Type targetItemType;
        private string propertyName;

        public PropertyObjectTypeResolver(Type targetItemType, string propertyName)
        {
            this.targetItemType = targetItemType;
            this.propertyName = propertyName;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var sourcePropertyValue = source.GetPublicInstancePropertyValue(propertyName);
            if (null != sourcePropertyValue)
            {
                var result = Activator.CreateInstance(this.targetItemType, true);
                PnPObjectsMapper.MapProperties(sourcePropertyValue, result, resolvers, recursive);

                return (result);
            }
            else
            {
                return (null);
            }
        }
    }
}
