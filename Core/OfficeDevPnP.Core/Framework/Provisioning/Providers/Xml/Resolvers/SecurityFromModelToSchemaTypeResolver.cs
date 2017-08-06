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
    /// Resolver for Security settings from model to schema
    /// </summary>
    internal class SecurityFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Object result = null;
            Boolean anySecurity = false;
            var security = source.GetPublicInstancePropertyValue("Security") as Model.ObjectSecurity;
            
            if (security != null)
            {
                var securityTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ObjectSecurity, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var securityType = Type.GetType(securityTypeName, true);
                var breakRoleInheritanceTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ObjectSecurityBreakRoleInheritance, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var breakRoleInheritanceType = Type.GetType(breakRoleInheritanceTypeName, true);
                var roleAssignmentTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.RoleAssignment, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var roleAssignmentType = Type.GetType(roleAssignmentTypeName, true);


                result = Activator.CreateInstance(securityType);
                var breakRoleInheritance = Activator.CreateInstance(breakRoleInheritanceType);

                breakRoleInheritance.SetPublicInstancePropertyValue("CopyRoleAssignments", security.CopyRoleAssignments);
                breakRoleInheritance.SetPublicInstancePropertyValue("ClearSubscopes", security.ClearSubscopes);

                var resolver = new CollectionFromModelToSchemaTypeResolver(roleAssignmentType);
                var roleAssignements = resolver.Resolve(security.RoleAssignments, null, true);
                breakRoleInheritance.SetPublicInstancePropertyValue("RoleAssignment", roleAssignements);

                anySecurity = (security.ClearSubscopes || security.CopyRoleAssignments || 
                    security.RoleAssignments != null && security.RoleAssignments.Count > 0);

                result.SetPublicInstancePropertyValue("BreakRoleInheritance", breakRoleInheritance);
            }

            return (anySecurity ? result : null);
        }
    }
}
