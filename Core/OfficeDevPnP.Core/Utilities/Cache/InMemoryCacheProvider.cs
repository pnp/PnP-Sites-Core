using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Cache
{
    public class InMemoryCacheProvider : ICacheProvider
    {
        private readonly Dictionary<string, object> _cacheStore = new Dictionary<string, object>();
        private readonly object _syncRoot = new object();

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