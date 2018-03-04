using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for Navigation Node from model to schema
    /// </summary>
    internal class NavigationNodeFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Array result = null;

            Object modelSource = source as Model.StructuralNavigation;
            if (modelSource == null)
            {
                modelSource = source as Model.NavigationNode;
            }
            
            if (modelSource != null)
            {
                Model.NavigationNodeCollection sourceNodes = modelSource.GetPublicInstancePropertyValue("NavigationNodes") as Model.NavigationNodeCollection;
                if (sourceNodes != null)
                {
                    var navigationNodeTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.NavigationNode, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var navigationNodeType = Type.GetType(navigationNodeTypeName, true);

                    result = Array.CreateInstance(navigationNodeType, sourceNodes.Count);

                    resolvers = new Dictionary<string, IResolver>();
                    resolvers.Add($"{navigationNodeType}.NavigationNode1", new NavigationNodeFromModelToSchemaTypeResolver());

                    for (Int32 c = 0; c < sourceNodes.Count; c++)
                    {
                        var targetNodeItem = Activator.CreateInstance(navigationNodeType);
                        PnPObjectsMapper.MapProperties(sourceNodes[c], targetNodeItem, resolvers, recursive);

                        result.SetValue(targetNodeItem, c);
                    }
                }
            }

            return (result);
        }
    }
}
