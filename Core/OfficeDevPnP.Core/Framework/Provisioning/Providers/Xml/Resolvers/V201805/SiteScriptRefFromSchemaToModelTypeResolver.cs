using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201805
{
    /// <summary>
    /// Resolves the Site Scripts reference for a Site Design from the Schema to the Model
    /// </summary>
    internal class SiteScriptRefFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public SiteScriptRefFromSchemaToModelTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            List<String> result = new List<String>();

            var siteScripts = source.GetPublicInstancePropertyValue("SiteScripts");
            if (null != siteScripts)
            {
                foreach (Object script in (IEnumerable<Object>)siteScripts)
                {
                    result.Add(script.GetPublicInstancePropertyValue("ID").ToString());
                }
            }

            return (result);
        }
    }
}
