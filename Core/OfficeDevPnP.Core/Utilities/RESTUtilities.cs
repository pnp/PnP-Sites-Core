using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    internal static class RESTUtilities
    {
        /// <summary>
        /// Sets the authentication cookie based upon either credentials currently used or if not set, the presence of any authentication cookies for the current context URL.
        /// </summary>
        /// <param name="handler"></param>
        /// <param name="context"></param>
        public static void SetAuthenticationCookies(this HttpClientHandler handler, ClientContext context)
        {
            if (context.Credentials != null)
            {
                handler.Credentials = context.Credentials;
                handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
            }
            else
            {
                var cookieString = CookieReader.GetCookie(context.Web.Url).Replace("; ", ",").Replace(";", ",");
                var authCookiesContainer = new System.Net.CookieContainer();
                // Get FedAuth and rtFa cookies issued by ADFS when accessing claims aware applications.
                // - or get the EdgeAccessCookie issued by the Web Application Proxy (WAP) when accessing non-claims aware applications (Kerberos).
                IEnumerable<string> authCookies = null;
                if (Regex.IsMatch(cookieString, "FedAuth", RegexOptions.IgnoreCase))
                {
                    authCookies = cookieString.Split(',').Where(c => c.StartsWith("FedAuth", StringComparison.InvariantCultureIgnoreCase) || c.StartsWith("rtFa", StringComparison.InvariantCultureIgnoreCase));
                }
                else if (Regex.IsMatch(cookieString, "EdgeAccessCookie", RegexOptions.IgnoreCase))
                {
                    authCookies = cookieString.Split(',').Where(c => c.StartsWith("EdgeAccessCookie", StringComparison.InvariantCultureIgnoreCase));
                }
                if (authCookies != null)
                {
                    authCookiesContainer.SetCookies(new Uri(context.Web.Url), string.Join(",", authCookies));
                }
                handler.CookieContainer = authCookiesContainer;
            }
        }
    }
}
