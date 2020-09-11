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
    /// Type resolver for SharingSettings from Model to Schema
    /// </summary>
    internal class SharingSettingsFromModelToSchema : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Object result = null;

            // Declare supporting types
            var sharingSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SharingSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var sharingSettingsType = Type.GetType(sharingSettingsTypeName, true);

            var settings = ((Model.ProvisioningTenant)source).SharingSettings;

            if (null != settings)
            {
                result = Activator.CreateInstance(sharingSettingsType);

                PnPObjectsMapper.MapProperties(settings, result, resolvers, recursive);

                if (settings.AllowedDomainList.Count > 0)
                {
                    var allowedDomainList = settings.AllowedDomainList.Aggregate(string.Empty, (acc, next) => acc += $",{next}");
                    allowedDomainList = allowedDomainList.Substring(1, allowedDomainList.Length - 1);
                    result.SetPublicInstancePropertyValue("AllowedDomainList", allowedDomainList);
                }
                if (settings.BlockedDomainList.Count > 0)
                {
                    var blockedDomainList = settings.BlockedDomainList.Aggregate(string.Empty, (acc, next) => acc += $",{next}");
                    blockedDomainList = blockedDomainList.Substring(1, blockedDomainList.Length - 1);
                    result.SetPublicInstancePropertyValue("BlockedDomainList", blockedDomainList);
                }
            }

            return (result);
        }
    }
}
