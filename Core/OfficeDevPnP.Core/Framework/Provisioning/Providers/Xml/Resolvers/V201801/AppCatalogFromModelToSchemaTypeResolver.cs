using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801
{
    internal class AppCatalogFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;


        public AppCatalogFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            // Try with the tenant-wide AppCatalog
            var tenant = source as Model.ProvisioningTenant;
            var appCatalog = tenant?.AppCatalog;

            if (null == appCatalog)
            {
                // If that one is missing, let's try with the local Site Collection App Catalog
                var alm = source as Model.ApplicationLifecycleManagement;
                appCatalog = alm.AppCatalog;
            }

            if (null != appCatalog)
            {
                var appCatalogPackageTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.AppCatalogPackage, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var appCatalogPackageType = Type.GetType(appCatalogPackageTypeName, true);

                resolvers = new Dictionary<string, IResolver>();
                resolvers.Add($"{appCatalogPackageType}.SkipFeatureDeploymentSpecified", new ExpressionValueResolver(() => true));

                var resolver = new CollectionFromModelToSchemaTypeResolver(appCatalogPackageType);
                result = resolver.Resolve(appCatalog.Packages, resolvers, true);
            }

            return (result);
        }
    }
}
