using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type to resolve a collection of DataRows for a ListInstance
    /// </summary>
    internal class ListInstanceDataRowValuesFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var dataRows = source?.GetType()?.GetProperty("DataRows", 
                    System.Reflection.BindingFlags.Instance | 
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase)?
                .GetValue(source);

            var result = new List<Model.DataRow>();

            if (dataRows != null)
            {
                foreach (var dr in (IEnumerable)dataRows)
                {
                    result.Add(new Model.DataRow { });
                }
            }

            return (null);
        }
    }
}
