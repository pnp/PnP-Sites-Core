using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201807
{
    internal class ClientSidePageSecurityFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;


        public ClientSidePageSecurityFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            // Try with the tenant-wide AppCatalog
            var page = source as Model.ClientSidePage;
            var security = page?.Security;

            // If we have security settings
            if (null != security && 
                security.RoleAssignments != null &&
                security.RoleAssignments.Count > 0)
            {
                // Map them to the output
                var securityTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ObjectSecurity, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var securityType = Type.GetType(securityTypeName, true);
                result = Activator.CreateInstance(securityType);

                PnPObjectsMapper.MapProperties(security, result, resolvers, recursive);
            }

            return (result);
        }
    }
}
