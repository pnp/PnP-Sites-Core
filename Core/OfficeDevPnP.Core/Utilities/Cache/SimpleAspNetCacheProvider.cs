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
        public T Get<T>(string cacheKey)
        {
            return (T)HttpRuntime.Cache.Get(cacheKey);
        }

        public void Put<T>(string cacheKey, T item)
        {
            HttpRuntime.Cache.Insert(cacheKey, item);
        }
    }
}