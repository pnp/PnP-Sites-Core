using Microsoft.SharePoint.Client;
using System;
using System.Reflection;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace OfficeDevPnP.Core
{
    /// <summary>
    /// Class that deals with PnPClientContext methods
    /// </summary>
    public class PnPClientContext : ClientContext
    {
        public int RetryCount { get; set; }
        public int Delay { get; set; }

        /// <summary>
        /// Generic property bag for the PnPClientContext
        /// </summary>
        public ConcurrentDictionary<string, object> PropertyBag { get; set; } = new ConcurrentDictionary<string, object>();

        /// <summary>
        /// Converts ClientContext into PnPClientContext
        /// </summary>
        /// <param name="clientContext">A SharePoint ClientContext for resource operations</param>
        /// <param name="retryCount">Maximum amount of retries before giving up</param>
        /// <param name="delay">Initial delay in milliseconds</param>
        /// <returns></returns>
        public static PnPClientContext ConvertFrom(ClientContext clientContext, int retryCount = 10, int delay = 500)
        {
            var pnpContext = new PnPClientContext(clientContext.Url,retryCount,delay);
            clientContext.Clone(pnpContext, new Uri(clientContext.Url));
            return pnpContext;
        }

        /// <summary>
        /// Creates a ClientContext allowing you to override the default retry and delay values of ExecuteQueryRetry
        /// </summary>
        /// <param name="url">A SharePoint site URL</param>
        /// <param name="retryCount">Maximum amount of retries before giving up</param>
        /// <param name="delay">Initial delay in milliseconds</param>
        public PnPClientContext(string url, int retryCount = 10, int delay = 500) : base(url)
        {
            RetryCount = retryCount;
            Delay = delay;
        }

        /// <summary>
        /// Creates a ClientContext allowing you to override the default retry and delay values of ExecuteQueryRetry
        /// </summary>
        /// <param name="uri">A SharePoint site/web full URL</param>
        /// <param name="retryCount">Maximum amount of retries before giving up</param>
        /// <param name="delay">Initial delay in milliseconds</param>
        public PnPClientContext(Uri uri, int retryCount = 10, int delay = 500) : base(uri)
        {
            RetryCount = retryCount;
            Delay = delay;
        }

       
    }
}
