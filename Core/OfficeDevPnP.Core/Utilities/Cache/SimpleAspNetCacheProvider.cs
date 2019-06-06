#if !NETSTANDARD2_0
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.Core.Utilities.Cache
{
    /// <summary>
    /// A Cache Provider implementation that uses the HttpRuntime Cache class as underlying implementation
    /// Is a suitable choice for implementation of standard IIS Distributed cache technique
    /// </summary>
    public class SimpleAspNetCacheProvider : ICacheProvider
    {
        /// <summary>
        /// Gets an item from the cache
        /// </summary>
        /// <typeparam name="T">Type of the object to get from cache</typeparam>
        /// <param name="cacheKey">Key of the object to get from cache</param>
        /// <returns>Default type value if not found, the object otherwise</returns>
        public T Get<T>(string cacheKey)
        {
            return (T)HttpRuntime.Cache.Get(cacheKey);
        }

        /// <summary>
        /// Stores an object in the cache. If the object exists it will be updated
        /// </summary>
        /// <typeparam name="T">Type of the object to store in the cache</typeparam>
        /// <param name="cacheKey">Key of the object to store in the cache</param>
        /// <param name="item">The actual object to store in the cache</param>
        public void Put<T>(string cacheKey, T item)
        {
            HttpRuntime.Cache.Insert(cacheKey, item);
        }
    }
}
#endif