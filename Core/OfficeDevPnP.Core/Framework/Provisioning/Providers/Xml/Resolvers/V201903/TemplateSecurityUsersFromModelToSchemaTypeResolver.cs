using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for Security from Model to Schema
    /// </summary>
    internal class TemplateSecurityUsersFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        private String collectionName;
        private String clearItemsPropertyName;

        public TemplateSecurityUsersFromModelToSchemaTypeResolver(String collectionName, String clearItemsPropertyName)
        {
            this.collectionName = collectionName;
            this.clearItemsPropertyName = clearItemsPropertyName;
        }

        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Object result = null;

            //source is either Security instance or SiteGroup instance
            var userCollection = source?.GetPublicInstancePropertyValue(this.collectionName);
            var clearItems = source?.GetPublicInstancePropertyValue(this.clearItemsPropertyName);

            if (null != userCollection && (((ICollection)userCollection).Count > 0 || (Boolean)clearItems))
            {
                var siteSecurityUsersCollectionTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.UsersList, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var siteSecurityUsersCollectionType = Type.GetType(siteSecurityUsersCollectionTypeName, true);
                var siteSecurityUserTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.User, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var siteSecurityUserType = Type.GetType(siteSecurityUserTypeName, true);

                result = Activator.CreateInstance(siteSecurityUsersCollectionType);

                var resolver = new CollectionFromModelToSchemaTypeResolver(siteSecurityUserType);
                result.SetPublicInstancePropertyValue("User",
                    resolver.Resolve(userCollection, resolvers, true));

                result.SetPublicInstancePropertyValue("ClearExistingItems", clearItems);
            }

            return (result);
        }
    }
}
