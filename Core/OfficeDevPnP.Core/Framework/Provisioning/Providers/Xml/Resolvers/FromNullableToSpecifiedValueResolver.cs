using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Decimal value into a Double
    /// </summary>
    internal class FromNullableToSpecifiedValueResolver<T> : IValueResolver
        where T : struct
    {
        private string propertySpecifiedName;

        public string Name => this.GetType().Name;

        public FromNullableToSpecifiedValueResolver(string propertySpecifiedName)
        {
            this.propertySpecifiedName = propertySpecifiedName;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            T res = default(T);
            if (sourceValue != null)
            {
                var nullable = sourceValue as Nullable<T>;
                if (nullable.HasValue)
                {
                    res = nullable.Value;
                    destination.SetPublicInstancePropertyValue(propertySpecifiedName, true);
                }
            }
            return res;
        }
    }
}
