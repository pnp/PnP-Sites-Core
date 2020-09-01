using System;
using System.Diagnostics;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that deals with cloning client context object, getting access token and validates server version
    /// </summary>
    public static partial class ClientContextExtensions
    {
        [Obsolete("Use GetRequestDigestAsync(ClientContext). This method will be removed in the November 2020 release.")]
        public static async Task<string> GetRequestDigest(this ClientContext context)
        {
            return await GetRequestDigestAsync(context);
        }

#if !ONPREMISES
        [Obsolete("Use HideTeamifyPromptAsync. The method will be removed in the November 2020 release.")]
        public static async Task<bool> HideTeamifyPrompt(this ClientContext clientContext)
        {
            return await HideTeamifyPromptAsync(clientContext);
        }
#endif
    }
}
