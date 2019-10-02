using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201909
{
    /// <summary>
    /// Type resolver for AlternateUICulture from Schema to Model
    /// </summary>
    internal class AlternateUICultureFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.AlternateUICulture>();

            var alternateUICultureItems = source.GetPublicInstancePropertyValue("AlternateUICultures");
            if (null != alternateUICultureItems)
            {
                foreach (var i in (IEnumerable)alternateUICultureItems)
                {
                    var targetItem = new Model.AlternateUICulture();
                    PnPObjectsMapper.MapProperties(i, targetItem, resolvers, recursive);
                    result.Add(targetItem);
                }
            }

            return (result);
        }
    }
}
