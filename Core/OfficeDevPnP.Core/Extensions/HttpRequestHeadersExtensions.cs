using System.Collections.Generic;
using System.Net.Http.Headers;

namespace OfficeDevPnP.Core.Extensions
{
    public static class HttpRequestHeadersExtensions
    {
        public static void AddDictionary(this HttpRequestHeaders header, Dictionary<string, string> additionalHeaders)
        {
            if (header == null || additionalHeaders == null)
            {
                return;
            }
            foreach (var additionalHeader in additionalHeaders)
            {
                header.Add(additionalHeader.Key, additionalHeader.Value);
            }
        }
    }
}