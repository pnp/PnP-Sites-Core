using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Decimal value into a Double
    /// </summary>
    internal class FromStringToGuidValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        public object Resolve(object source, object destination, object sourceValue)
        {
            return sourceValue != null ? (Guid.Parse((String)sourceValue)) : default(Guid);
        }
    }
}
