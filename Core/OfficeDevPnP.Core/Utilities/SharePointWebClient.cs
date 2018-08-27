using Microsoft.SharePoint.Client;
using System;
using System.Net;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// WebClient wrapper that automate getting credentials from a client context
    /// </summary>
    public class SharePointWebClient : WebClient
    {
        private readonly ClientRuntimeContext _ctx;

        /// <summary>
        /// Initializes a new SharePointWebClient using a given SharePoint client context
        /// </summary>
        /// <param name="ctx"></param>
        public SharePointWebClient(ClientRuntimeContext ctx)
        {
            if (ctx == null) throw new ArgumentNullException(nameof(ctx));

            _ctx = ctx;
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            var req = base.GetWebRequest(address);
            ClientContextExtensions.SetupWebRequest(_ctx, (HttpWebRequest)req);
            return req;
        }
    }
}