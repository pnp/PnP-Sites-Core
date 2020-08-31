using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
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
#if !NETSTANDARD2_0
            if (context.Credentials is SharePointOnlineCredentials spCred)
            {
                handler.Credentials = context.Credentials;
                handler.CookieContainer.SetCookies(new Uri(context.Web.Url), spCred.GetAuthenticationCookie(new Uri(context.Web.Url)));
            }
            else 
#endif            
            if (context.Credentials == null)
            {
                var cookieString = CookieReader.GetCookie(context.Web.Url)?.Replace("; ", ",")?.Replace(";", ",");
                if(cookieString == null)
                {
                    return;
                }
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

        /// <summary>
        /// Executes a REST Get request. 
        /// </summary>
        /// <param name="web">The current web to execute the request against</param>
        /// <param name="endpoint">The full endpoint url, exluding the URL of the web, e.g. /_api/web/lists</param>
        /// <returns></returns>
        internal static async Task<string> ExecuteGetAsync(this Web web, string endpoint)
        {
            string returnObject = null;
            var accessToken = web.Context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                web.EnsureProperty(w => w.Url);

                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(web.Context as ClientContext);
                }


                using (var httpClient = new PnPHttpProvider(handler))
                {
                    var requestUrl = $"{web.Url}{endpoint}";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=nometadata");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (web.Context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }

                    var requestDigest = await (web.Context as ClientContext).GetRequestDigestAsync().ConfigureAwait(false);
                    request.Headers.Add("X-RequestDigest", requestDigest);

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            try
                            {

                                returnObject = responseString;
                        
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return await Task.Run(() => returnObject);
        }

        internal static async Task<string> ExecutePostAsync(this Web web, string endpoint, string payload)
        {
            string returnObject = null;
            var accessToken = web.Context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                web.EnsureProperty(w => w.Url);

                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(web.Context as ClientContext);
                }


                using (var httpClient = new PnPHttpProvider(handler))
                {
                    var requestUrl = $"{web.Url}{endpoint}";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=nometadata");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (web.Context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }
                    var requestDigest = await (web.Context as ClientContext).GetRequestDigestAsync().ConfigureAwait(false);
                    request.Headers.Add("X-RequestDigest", requestDigest);

                    if (!string.IsNullOrEmpty(payload))
                    {
                        ////var jsonBody = JsonConvert.SerializeObject(postObject);
                        var requestBody = new StringContent(payload);
                        MediaTypeHeaderValue sharePointJsonMediaType;
                        MediaTypeHeaderValue.TryParse("application/json;odata=nometadata;charset=utf-8", out sharePointJsonMediaType);
                        requestBody.Headers.ContentType = sharePointJsonMediaType;
                        request.Content = requestBody;
                    }
                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        if (responseString != null)
                        {
                            try
                            {

                                returnObject = responseString;

                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return await Task.Run(() => returnObject);
        }
    }
}
