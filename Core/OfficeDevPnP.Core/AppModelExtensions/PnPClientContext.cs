using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core
{
    public class PnPClientContext : ClientContext
    {
        public int RetryCount { get; protected set; }
        public int Delay { get; protected set; }

        /// <summary>
        /// Creates a ClientContext allowing you to override the default retry and delay values of ExecuteQueryRetry
        /// </summary>
        /// <param name="url"></param>
        /// <param name="retryCount"></param>
        /// <param name="delay"></param>
        public PnPClientContext(string url, int retryCount = 10, int delay = 500) : base(url)
        {
            this.RetryCount = retryCount;
            this.Delay = delay;
        }

        /// <summary>
        /// Creates a ClientContext allowing you to override the default retry and delay values of ExecuteQueryRetry
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="retryCount"></param>
        /// <param name="delay"></param>
        public PnPClientContext(Uri uri, int retryCount = 10, int delay = 500) : base (uri)
        {
            this.RetryCount = retryCount;
            this.Delay = delay;
        }
    }
}
