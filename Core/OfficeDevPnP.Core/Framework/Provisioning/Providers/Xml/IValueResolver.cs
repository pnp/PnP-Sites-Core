using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Handles custom value resolving rules for PnPObjectsMapper
    /// </summary>
    internal interface IValueResolver : IResolver
    {
        /// <summary>
        /// Resolves a source value into a result
        /// </summary>
        /// <param name="source">The full source object to resolve</param>
        /// <param name="destination">The full destination object to resolve</param>
        /// <param name="sourceValue">The source value to resolve</param>
        /// <returns>The resolved value</returns>
        Object Resolve(Object source, Object destination, Object sourceValue);
    }
}
