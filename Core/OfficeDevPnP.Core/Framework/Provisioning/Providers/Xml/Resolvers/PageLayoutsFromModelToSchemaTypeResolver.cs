using OfficeDevPnP.Core.Framework.Provisioning.Model;
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
    internal class PageLayoutsFromModelToSchemaTypeResolver: ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            object result = null;
            if (source != null)
            {
                var layouts = ((Publishing)source).PageLayouts;
                if (layouts != null && layouts.Count > 0)
                {
                    var publishingPageLayoutsType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PublishingPageLayouts, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                    var publishingPageLayoutType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.PublishingPageLayoutsPageLayout, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
                    var target = Activator.CreateInstance(publishingPageLayoutsType, true);
                    var defaultLayout = layouts.Any(p => p.IsDefault) ? layouts.Last(p => p.IsDefault).Path : null;
                    target.GetPublicInstanceProperty("Default").SetValue(target, defaultLayout);

                    var targetLayouts = PnPObjectsMapper.MapObjects(layouts, new CollectionFromModelToSchemaTypeResolver(publishingPageLayoutType), null, true);
                    target.GetPublicInstanceProperty("PageLayout").SetValue(target, targetLayouts);
                    result = target;
                }
            }
            return result;
        }
    }
}
