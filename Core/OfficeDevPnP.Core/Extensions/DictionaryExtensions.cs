using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Extensions
{
    /// <summary>
    /// Extension type for Dictionaries
    /// </summary>
    public static class DictionaryExtensions
    {
        public static void AddRange<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, IDictionary<TKey, TValue> range)
        {
            foreach (var item in range)
            {
                dictionary.Add(item.Key, item.Value);
            }
        }
    }
}
