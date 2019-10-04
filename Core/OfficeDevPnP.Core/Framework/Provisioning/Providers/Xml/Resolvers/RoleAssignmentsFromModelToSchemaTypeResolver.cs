using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a collection type from Domain Model to Schema
    /// </summary>
    internal class RoleAssignmentsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public RoleAssignmentsFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var baseNamespace = PnPSerializationScope.Current?.BaseSchemaNamespace;
            var breakRoleInheritanceTypeName = $"{baseNamespace}.ObjectSecurityBreakRoleInheritance, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var breakRoleInheritanceType = Type.GetType(breakRoleInheritanceTypeName, true);
            var roleAssignmentTypeName = $"{baseNamespace}.RoleAssignment, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var roleAssignmentType = Type.GetType(roleAssignmentTypeName, true);

            var security = (ObjectSecurity)source;

            // If we have any of the flags configured, or any role assignment, then we process the item
            if (security.ClearSubscopes || security.CopyRoleAssignments || 
                (security.RoleAssignments != null && security.RoleAssignments.Count > 0))
            {
                var breakRoleInheritance = Activator.CreateInstance(breakRoleInheritanceType, true);

                PnPObjectsMapper.MapProperties(source, breakRoleInheritance, recursive: true);

                if (security.RoleAssignments != null)
                {
                    var roleAssignment = PnPObjectsMapper.MapObjects(security.RoleAssignments,
                        new CollectionFromModelToSchemaTypeResolver(roleAssignmentType), null, true);
                    breakRoleInheritance.GetPublicInstanceProperty("RoleAssignment").SetValue(breakRoleInheritance, roleAssignment);
                }

                return breakRoleInheritance;
            }
            else
            {
                return null;
            }
        }
    }
}
