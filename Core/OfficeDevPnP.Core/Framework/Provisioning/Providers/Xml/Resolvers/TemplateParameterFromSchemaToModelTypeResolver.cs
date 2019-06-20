using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Template Parameter type from Schema to Domain Model
    /// </summary>
    internal class TemplateParameterFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new Dictionary<String, String>();

            if (null != source)
            {
                foreach (var l in (IEnumerable)source)
                {
                    // TODO: Extract key and text
                    var key = (String)l.GetType().GetProperty("Key").GetValue(l);
                    var value = ((String[])l.GetType().GetProperty("Text").GetValue(l))?
                        .Aggregate(String.Empty, (s, v) => s += v);
                    result.Add(key, value);
                }
            }

            return (result);
        }
    }
}
