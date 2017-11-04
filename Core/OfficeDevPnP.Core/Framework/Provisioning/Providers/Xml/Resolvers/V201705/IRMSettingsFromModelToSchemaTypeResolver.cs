using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    /// <summary>
    /// Resolves the IRM Settings of a List from Domain Model to Schema
    /// </summary>
    internal class IRMSettingsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => true;


        public IRMSettingsFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var list = source as Model.ListInstance;
            var irmSettings = list?.IRMSettings;

            var irmSettingsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.IRMSettings, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var irmSettingsType = Type.GetType(irmSettingsTypeName, true);

            Object result = null;

            if (null != irmSettings)
            {
                result = Activator.CreateInstance(irmSettingsType);
                PnPObjectsMapper.MapProperties(irmSettings, result, resolvers, recursive);

                result.SetPublicInstancePropertyValue("DocumentAccessExpireDaysSpecified", true);
                result.SetPublicInstancePropertyValue("LicenseCacheExpireDaysSpecified", true);
            }

            return (result);
        }
    }
}
