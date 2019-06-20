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
    internal class FromStringToEnumValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        private Type _targetItemType;

        public FromStringToEnumValueResolver(Type targetItemType)
        {
            _targetItemType = targetItemType;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            var s = sourceValue != null ? sourceValue.ToString() : null;
            return !string.IsNullOrEmpty(s) ? Enum.Parse(_targetItemType, s, true) : 0;
        }
    }
}
