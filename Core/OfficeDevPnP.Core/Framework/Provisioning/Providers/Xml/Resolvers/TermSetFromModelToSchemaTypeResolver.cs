using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a TermSet collection type from Domain Model to Schema
    /// </summary>
    internal class TermSetFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;


        private Type _targetItemType;

        public TermSetFromModelToSchemaTypeResolver()
        {
            this._targetItemType = Type.GetType($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TermSet, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);
        }

        //TermSet should never be null in schema
        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = Array.CreateInstance(this._targetItemType, 0);
            if (null != source)
            {
                var termGroup = (TermGroup)source;

                if (termGroup.TermSets != null && termGroup.TermSets.Count > 0)
                {
                    result = Array.CreateInstance(this._targetItemType, termGroup.TermSets.Count);
                    var index = 0;
                    foreach (var i in termGroup.TermSets)
                    {
                        var targetItem = Activator.CreateInstance(this._targetItemType, true);
                        PnPObjectsMapper.MapProperties(i, targetItem, resolvers, recursive);
                        result.SetValue(targetItem, index);
                        index++;
                    }
                }
            }
            return (result);
        }
    }
}
