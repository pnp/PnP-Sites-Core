using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a list of Views from Schema to Domain Model
    /// </summary>
    internal class PageLayoutsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new List<Model.PageLayout>();

            var layouts = source.GetPublicInstancePropertyValue("PageLayouts");
            if (layouts != null)
            {
                var defaultLayout = layouts.GetPublicInstancePropertyValue("Default");
                var layoutCollection = layouts.GetPublicInstancePropertyValue("PageLayout");
                if (layoutCollection != null)
                {
                    foreach(var layout in (IEnumerable)layoutCollection)
                    {
                        var path = layout.GetPublicInstancePropertyValue("Path");
                        result.Add(new Model.PageLayout() { Path = path?.ToString(), IsDefault = string.Equals(path, defaultLayout) });
                    }
                }
             }
            return (result);
        }
    }
}
