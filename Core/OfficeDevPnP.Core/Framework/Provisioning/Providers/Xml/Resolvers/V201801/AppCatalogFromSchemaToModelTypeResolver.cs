using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801
{
    /// <summary>
    /// Resolves the AppCatalog settings at the Tenant level from the Schema to the Model
    /// </summary>

    internal class AppCatalogFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public AppCatalogFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new AppCatalog();

            var appCatalogPackages = source.GetPublicInstancePropertyValue("AppCatalog");

            if (null != appCatalogPackages)
            {
                foreach (var p in ((IEnumerable)appCatalogPackages))
                {
                    var targetItem = new Model.Package();
                    PnPObjectsMapper.MapProperties(p, targetItem, resolvers, recursive);
                    result.Packages.Add(targetItem);
                }
            }

            return (result);
        }
    }
}
