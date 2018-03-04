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
    /// Resolves a collection type from Domain Model to Schema
    /// </summary>
    internal class RoleAssigmentsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

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
