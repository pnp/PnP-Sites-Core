using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    /// <summary>
    /// Resolves the IRM Settings of a List from Schema to Domain Model
    /// </summary>
    internal class IRMSettingsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => true;


        public IRMSettingsFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            var irmSettings = source.GetPublicInstancePropertyValue("IRMSettings");
            if (null != irmSettings)
            {
                result = new Model.IRMSettings();
                PnPObjectsMapper.MapProperties(irmSettings, result, resolvers, recursive);
            }

            return (result);
        }
    }
}
