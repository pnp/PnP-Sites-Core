using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801
{
    /// <summary>
    /// Resolves the CDN settings at the Tenant level from the Schema to the Model
    /// </summary>
    internal class CdnFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public CdnFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            ContentDeliveryNetwork result = null;

            var cdnSettings = source.GetPublicInstancePropertyValue("ContentDeliveryNetwork");

            if (null != cdnSettings)
            {
                var publicCdn = cdnSettings.GetPublicInstancePropertyValue("Public");
                CdnSettings publicCdnSettings = null;

                if (null != publicCdn)
                {
                    publicCdnSettings = new Model.CdnSettings();
                    PnPObjectsMapper.MapProperties(publicCdn, publicCdnSettings, resolvers, recursive);
                }

                var privateCdn = cdnSettings.GetPublicInstancePropertyValue("Private");
                CdnSettings privateCdnSettings = null;

                if (null != privateCdn)
                {
                    privateCdnSettings = new Model.CdnSettings();
                    PnPObjectsMapper.MapProperties(privateCdn, privateCdnSettings, resolvers, recursive);
                }

                result = new ContentDeliveryNetwork(publicCdnSettings, privateCdnSettings);
            }

            return (result);
        }
    }
}
