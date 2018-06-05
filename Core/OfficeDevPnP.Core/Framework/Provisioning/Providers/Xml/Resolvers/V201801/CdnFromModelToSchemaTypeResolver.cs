using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801
{
    /// <summary>
    /// Resolves the CDN settings at the Tenant level from the Model to the Schema
    /// </summary>
    internal class CdnFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;


        public CdnFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            var tenant = source as Model.ProvisioningTenant;
            var cdn = tenant?.ContentDeliveryNetwork;

            if (null != cdn)
            {
                var cdnTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ContentDeliveryNetwork, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var cdnType = Type.GetType(cdnTypeName, true);
                var cdnSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CdnSetting, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var cdnSettingsType = Type.GetType(cdnSettingsTypeName, true);

                result = Activator.CreateInstance(cdnType);

                if (cdn.PublicCdn != null)
                {
                    Object publicCdn = Activator.CreateInstance(cdnSettingsType);
                    PnPObjectsMapper.MapProperties(cdn.PublicCdn, publicCdn, null, true);
                    result.SetPublicInstancePropertyValue("Public", publicCdn);
                }

                if (cdn.PrivateCdn != null)
                {
                    Object privateCdn = Activator.CreateInstance(cdnSettingsType);
                    PnPObjectsMapper.MapProperties(cdn.PrivateCdn, privateCdn, null, true);
                    result.SetPublicInstancePropertyValue("Private", privateCdn);
                }
            }

            return (result);
        }
    }
}
