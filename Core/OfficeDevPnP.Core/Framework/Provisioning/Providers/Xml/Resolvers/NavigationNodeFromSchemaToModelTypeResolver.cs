using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for Navigation Node from schema to model
    /// </summary>
    internal class NavigationNodeFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.NavigationNode>();

            var nodes = source.GetPublicInstancePropertyValue("NavigationNode");
            if (null == nodes)
            {
                nodes = source.GetPublicInstancePropertyValue("NavigationNode1");
            }

            resolvers = new Dictionary<string, IResolver>();
            resolvers.Add($"{typeof(Model.NavigationNode).FullName}.NavigationNodes", new NavigationNodeFromSchemaToModelTypeResolver());

            if (null != nodes)
            {
                foreach (var f in ((IEnumerable)nodes))
                {
                    var targetItem = new Model.NavigationNode();
                    PnPObjectsMapper.MapProperties(f, targetItem, resolvers, recursive);
                    result.Add(targetItem);
                }
            }

            return (result);
        }
    }
}
