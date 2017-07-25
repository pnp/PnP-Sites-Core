using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class FoldersFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new List<Model.Folder>();

            var folders = source.GetPublicInstancePropertyValue("Folders");
            if (null == folders)
            {
                folders = source.GetPublicInstancePropertyValue("Folder1");
            }

            resolvers = new Dictionary<string, IResolver>();
            resolvers.Add($"{typeof(Model.Folder).FullName}.Folders", new FoldersFromSchemaToModelTypeResolver());
            resolvers.Add($"{typeof(Model.Folder).FullName}.Security", new SecurityFromSchemaToModelTypeResolver());

            if (null != folders)
            {
                foreach (var f in ((IEnumerable)folders))
                {
                    var targetItem = new Model.Folder();
                    PnPObjectsMapper.MapProperties(f, targetItem, resolvers, recursive);
                    result.Add(targetItem);
                }
            }

            return (result);
        }
    }
}
