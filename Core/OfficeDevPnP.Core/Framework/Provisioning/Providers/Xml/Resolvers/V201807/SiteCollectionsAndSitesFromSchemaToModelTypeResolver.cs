using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201807
{
    /// <summary>
    /// Allows resolving specific SiteCollection and SubSite types
    /// </summary>
    internal class SiteCollectionsAndSitesFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        private Type _targetItemType;

        public SiteCollectionsAndSitesFromSchemaToModelTypeResolver(Type targetItemType)
        {
            this._targetItemType = targetItemType;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var itemType = typeof(List<>);
            var resultType = itemType.MakeGenericType(new Type[] { this._targetItemType });
            IList result = (IList)Activator.CreateInstance(resultType);

            // Define the specific source schema types
            var communicationSiteTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CommunicationSite, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var communicationSiteType = Type.GetType(communicationSiteTypeName, true);
            var teamSiteTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSite, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamSiteType = Type.GetType(teamSiteTypeName, true);
            var teamSiteNoGroupTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSiteNoGroup, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamSiteNoGroupType = Type.GetType(teamSiteNoGroupTypeName, true);
            var teamSubSiteNoGroupTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSubSiteNoGroup, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamSubSiteNoGroupType = Type.GetType(teamSubSiteNoGroupTypeName, true);

            // Retrieve the source collection
            var sourceCollection = source.GetPublicInstancePropertyValue("SiteCollections");
            if (sourceCollection == null)
            {
                sourceCollection = source.GetPublicInstancePropertyValue("Sites");
            }

            if (null != sourceCollection)
            {
                foreach (var i in (IEnumerable)sourceCollection)
                {
                    Object targetItem = null;

                    if (i.GetType().Name == communicationSiteType.Name)
                    {
                        targetItem = new Model.CommunicationSiteCollection();
                    }
                    else if (i.GetType().Name == teamSiteType.Name)
                    {
                        targetItem = new Model.TeamSiteCollection();
                    }
                    else if (i.GetType().Name == teamSiteNoGroupType.Name)
                    {
                        targetItem = new Model.TeamNoGroupSiteCollection();
                    }
                    else if (i.GetType().Name == teamSubSiteNoGroupType.Name)
                    {
                        targetItem = new Model.TeamNoGroupSubSite();
                    }

                    PnPObjectsMapper.MapProperties(i, targetItem, resolvers, recursive);

                    if (targetItem!= null)
                    {
                        result.Add(targetItem);
                    }
                }
            }

            return (result);
        }
    }
}
