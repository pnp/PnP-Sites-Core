using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Cache
{
    /// <summary>
    /// The interface of cache Provider
    /// </summary>
    public interface ICacheProvider
    {
        /// <summary>
        /// Puts an entry in cache
        /// </summary>
        /// <typeparam name="T">The type of item to cache</typeparam>
        /// <param name="cacheKey">The key for the cached item</param>
        /// <param name="item">The item to put into cache</param>
        void Put<T>(string cacheKey, T item);

        /// <summary>
        /// Returns an item in cache keyed by the cacheKey parameter
        /// </summary>
        /// <typeparam name="T">The expected type of cached item</typeparam>
        /// <param name="cacheKey">The key for the cached item</param>
        /// <returns>The item retrieved from the cache</returns>
        T Get<T>(string cacheKey);
    }
}