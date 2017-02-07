using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a type from Schema to Domain Model
    /// </summary>
    internal class CollectionFromSchemaToModelTypeResolver : ITypeResolver
    {
        private Type _targetItemType;

        public CollectionFromSchemaToModelTypeResolver(Type targetItemType)
        {
            this._targetItemType = targetItemType;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null)
        {
            var itemType = typeof(List<>);
            var resultType = itemType.MakeGenericType(new Type[] { this._targetItemType });
            IList result = (IList)Activator.CreateInstance(resultType);

            if (null != source)
            {
                foreach (var i in (IEnumerable)source)
                {
                    var targetItem = Activator.CreateInstance(this._targetItemType);
                    PnPObjectsMapper.MapProperties(i, targetItem, resolvers);
                    result.Add(targetItem);
                }
            }

            return (result);
        }
    }
}
