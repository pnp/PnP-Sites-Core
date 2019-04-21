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
    internal class SiteCollectionsAndSitesFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        private Type _targetItemType;

        public SiteCollectionsAndSitesFromModelToSchemaTypeResolver(Type targetItemType)
        {
            this._targetItemType = targetItemType;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
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

            Array resultArray = null;

            if (null != sourceCollection)
            {
                var resultType = this._targetItemType.MakeArrayType();
                resultArray = (Array)Activator.CreateInstance(resultType, ((IList)sourceCollection).Count);
                var i = 0;

                foreach (var sourceItem in (IEnumerable)sourceCollection)
                {
                    Object targetItem = null;

                    switch (sourceItem)
                    {
                        case CommunicationSiteCollection cs:
                            targetItem = Activator.CreateInstance(communicationSiteType);
                            break;
                        case TeamSiteCollection ts:
                            targetItem = Activator.CreateInstance(teamSiteType);
                            break;
                        case TeamNoGroupSiteCollection tngs:
                            targetItem = Activator.CreateInstance(teamSiteNoGroupType);
                            break;
                        case TeamNoGroupSubSite tngss:
                            targetItem = Activator.CreateInstance(teamSubSiteNoGroupType);
                            break;
                    }

                    PnPObjectsMapper.MapProperties(sourceItem, targetItem, resolvers, recursive);

                    if (targetItem != null)
                    {
                        resultArray.SetValue(targetItem, i++);
                    }
                }
            }

            return (resultArray.Length > 0 ? resultArray : null);
        }
    }
}
