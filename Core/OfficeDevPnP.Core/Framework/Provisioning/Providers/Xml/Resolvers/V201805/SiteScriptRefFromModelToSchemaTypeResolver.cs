using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201805
{
    /// <summary>
    /// Resolves the Site Scripts reference for a Site Design from the Model to the Schema
    /// </summary>
    internal class SiteScriptRefFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;


        public SiteScriptRefFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Array result = null;

            var scripts = (source as Model.SiteDesign)?.SiteScripts;
            var siteDesignScriptsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteDesignsSiteDesignSiteScriptRef, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var siteDesignScriptsType= Type.GetType(siteDesignScriptsTypeName, true);

            if (null != scripts)
            {
                result = Array.CreateInstance(siteDesignScriptsType, scripts.Count);

                for (Int32 c = 0; c < scripts.Count; c++)
                {
                    var scriptRef = Activator.CreateInstance(siteDesignScriptsType);
                    scriptRef.SetPublicInstancePropertyValue("ID", scripts[c]);

                    result.SetValue(scriptRef, c);
                }
            }

            return (result);
        }
    }
}
