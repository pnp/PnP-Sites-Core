using OfficeDevPnP.Core.Framework.Provisioning.Model;
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
    internal class PropertyCollectionTypeResolver<T> : PropertyObjectTypeResolver
    {
        public PropertyCollectionTypeResolver(Expression<Func<T, object>> exp, Func<object, object> sourceValueSelector = null) : 
            base(GetPropertyType(exp), GetPropertyName(exp), sourceValueSelector)
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
    internal class RoleAssigmentsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }


        public RoleAssigmentsFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            List<RoleAssignment> res = new List<RoleAssignment>();
            var sourceValue = source.GetPublicInstancePropertyValue("RoleAssignment");
            if(sourceValue != null)
            {
                res = PnPObjectsMapper.MapObjects(sourceValue, new CollectionFromSchemaToModelTypeResolver(typeof(RoleAssignment)), null, true) as List<RoleAssignment>;
            }
            return res;
        }
    }
}
