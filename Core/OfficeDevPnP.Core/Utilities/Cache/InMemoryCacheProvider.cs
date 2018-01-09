using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Cache
{
    /// <summary>
    /// Simple in memory cache provider
    /// </summary>
    public class InMemoryCacheProvider : ICacheProvider
    {
        private readonly Dictionary<string, object> _cacheStore = new Dictionary<string, object>();
        private readonly object _syncRoot = new object();

        /// <summary>
        /// Gets an item from the cache
        /// </summary>
        /// <typeparam name="T">Type of the object to get from cache</typeparam>
        /// <param name="cacheKey">Key of the object to get from cache</param>
        /// <returns>Default type value if not found, the object otherwise</returns>
        public T Get<T>(string cacheKey)
        {
            lock (_syncRoot)
            {
                if (!_cacheStore.ContainsKey(cacheKey))
                {
                    return default(T);
                }
                return (T)_cacheStore[cacheKey];
            }
        }

        /// <summary>
        /// Stores an object in the cache. If the object exists it will be updated
        /// </summary>
        /// <typeparam name="T">Type of the object to store in the cache</typeparam>
        /// <param name="cacheKey">Key of the object to store in the cache</param>
        /// <param name="item">The actual object to store in the cache</param>
        public void Put<T>(string cacheKey, T item)
        {
            lock (_syncRoot)
            {
                if (!_cacheStore.ContainsKey(cacheKey))
                {
                    _cacheStore.Add(cacheKey, item);
                }
                else
                {
                    _cacheStore[cacheKey] = item;
                }
            }
        }
    }
}