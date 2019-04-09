using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves the AAD Users from the Model to the Schema
    /// </summary>
    internal class AADUsersFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            // Try with the tenant-wide AppCatalog
            var aad = source as Model.AzureActiveDirectory.ProvisioningAzureActiveDirectory;
            var users = aad?.Users;

            if (null != users)
            {
                var aadUsersTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.AADUsersUser, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var aadUsersType = Type.GetType(aadUsersTypeName, true);

                var resolver = new CollectionFromModelToSchemaTypeResolver(aadUsersType);
                result = resolver.Resolve(users, resolvers, true);
            }

            return (result);
        }
    }
}
