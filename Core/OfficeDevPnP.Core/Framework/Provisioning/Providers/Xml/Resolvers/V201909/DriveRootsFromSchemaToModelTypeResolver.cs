using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for DriveRoots from Schema to Model
    /// </summary>
    internal class DriveRootsFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.Drive.DriveRoot>();

            var driveRootTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveRoot, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var driveRootType = Type.GetType(driveRootTypeName, true);

            if (null != source)
            {
                foreach (var dr in ((IEnumerable)source))
                {
                    if (driveRootType.IsInstanceOfType(dr))
                    {
                        var targetItem = new Model.Drive.DriveRoot();
                        PnPObjectsMapper.MapProperties(dr, targetItem, resolvers, recursive);
                        result.Add(targetItem);
                    }
                }
            }

            return (result);
        }
    }
}
