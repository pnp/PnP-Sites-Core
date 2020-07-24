using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V202002
{
    /// <summary>
    /// Type resolver for SharingSettings from Schema to Model
    /// </summary>
    internal class SharingSettingsFromSchemaToModel : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new Model.SharingSettings();
            var settings = source.GetPublicInstancePropertyValue("SharingSettings");

            if (null != settings)
            {
                PnPObjectsMapper.MapProperties(settings, result, resolvers, recursive);

                var allowedDomainList = (String)settings.GetPublicInstancePropertyValue("AllowedDomainList");
                if (!String.IsNullOrEmpty(allowedDomainList))
                {
                    result.AllowedDomainList.AddRange(allowedDomainList.Split(','));
                }
                var blockedDomainList = (String)settings.GetPublicInstancePropertyValue("BlockedDomainList");
                if (!String.IsNullOrEmpty(blockedDomainList))
                {
                    result.BlockedDomainList.AddRange(blockedDomainList.Split(','));
                }
            }

            return (result);
        }
    }
}
