using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolver for Security settings from schema to model
    /// </summary>
    internal class SecurityFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new ObjectSecurity();

            // Use Reflection to get the Security property
            var security = source.GetPublicInstancePropertyValue("Security");

            if (null != security)
            {
                var breakRoleInheritance = security.GetPublicInstancePropertyValue("BreakRoleInheritance");

                if (null != breakRoleInheritance)
                {
                    var copyRoleAssignments = breakRoleInheritance.GetPublicInstancePropertyValue("CopyRoleAssignments");
                    if (null != copyRoleAssignments)
                    {
                        result.CopyRoleAssignments = (Boolean)copyRoleAssignments;
                    }

                    var clearSubscopes = breakRoleInheritance.GetPublicInstancePropertyValue("ClearSubscopes");
                    if (null != clearSubscopes)
                    {
                        result.ClearSubscopes = (Boolean)clearSubscopes;
                    }

                    var roleAssignments = breakRoleInheritance.GetPublicInstancePropertyValue("RoleAssignment");
                    result.RoleAssignments.AddRange(
                        PnPObjectsMapper.MapObjects<ListInstance>(roleAssignments,
                                new CollectionFromSchemaToModelTypeResolver(typeof(RoleAssignment)),
                                null,
                                recursive: true)
                                as IEnumerable<RoleAssignment>);                    
                }
            }

            return (result);
        }
    }
}
