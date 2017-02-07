using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Handles custom type resolving rules for PnPObjectsMapper
    /// </summary>
    public interface ITypeResolver : IResolver
    {
        /// <summary>
        /// Resolves a source type into a result
        /// </summary>
        /// <param name="source">The full source object to resolve</param>
        Object Resolve(Object source, Dictionary<String, IResolver> resolvers = null);
    }
}
