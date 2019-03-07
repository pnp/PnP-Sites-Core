using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Utilities.Async;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Themes
{
#if !ONPREMISES
    /// <summary>
    /// Extension class for the Web object useful to apply custom Themes
    /// </summary>
    public static class ThemeManager
    {
        /// <summary>
        /// Extension method to apply a Theme to a target web
        /// </summary>
        /// <param name="web"></param>
        /// <param name="jsonTheme"></param>
        /// <param name="themeName"></param>
        /// <returns></returns>
        public static Boolean ApplyTheme(this Web web, String jsonTheme, String themeName = null)
        {
            return Task.Run(() => ApplyThemeAsync(web, jsonTheme, themeName)).GetAwaiter().GetResult();
        }

        public static async Task<Boolean> ApplyThemeAsync(Web web, String jsonTheme, String themeName = null)
        {
            if (web == null)
            {
                throw new ArgumentException(nameof(web));
            }
            if (string.IsNullOrEmpty(jsonTheme))
            {
                throw new ArgumentException(nameof(jsonTheme));
            }

            await new SynchronizationContextRemover();

            return await BaseRequest((ClientContext)web.Context, ThemeAction.ApplyTheme,
                new
                {
                    name = themeName ?? String.Empty,
                    themeJson = JsonConvert.SerializeObject(new
                    {
                        palette = JsonConvert.DeserializeObject(jsonTheme)
                    }),
                });
        }

        /// <summary>
        /// Internal private method to process an HTTP request toward the ThemeManager REST APIs
        /// </summary>
        /// <param name="context">The current ClientContext of CSOM</param>
        /// <param name="action">The action to perform</param>
        /// <param name="postObject">The body of the request</param>
        /// <param name="accessToken">An optional Access Token for OAuth authorization</param>
        /// <returns>A boolean declaring whether the operation was successful</returns>
        private static async Task<bool> BaseRequest(ClientContext context, ThemeAction action, Object postObject, String accessToken = null)
        {
            var result = false;

            // If we don't have the access token
            if (String.IsNullOrEmpty(accessToken))
            {
                // Try to get one from the current context
                accessToken = context.GetAccessToken();
            }

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // We're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    // Reference here: https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-rest-api
                    string requestUrl = $"{context.Web.Url}/_api/thememanager/{action.ToString()}";

                    // Always make a POST request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("ACCEPT", "application/json; odata.metadata=minimal");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());
                    request.Headers.Add("ODATA-VERSION", "4.0");

                    if (postObject != null)
                    {
                        var jsonBody = JsonConvert.SerializeObject(postObject);
                        var requestBody = new StringContent(jsonBody);
                        MediaTypeHeaderValue sharePointJsonMediaType;
                        MediaTypeHeaderValue.TryParse("application/json;charset=utf-8", out sharePointJsonMediaType);
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
                                var responseJson = JObject.Parse(responseString);
                                result = true;
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
            return await Task.Run(() => result);
        }
    }

    /// <summary>
    /// Available actions for ThemeManager REST APIs
    /// </summary>
    public enum ThemeAction
    {
        /// <summary>
        /// Add a new Tenant Theme
        /// </summary>
        AddTenantTheme,
        /// <summary>
        /// Delete an existing Tenant Theme
        /// </summary>
        DeleteTenantTheme,
        /// <summary>
        /// Get the whole list of Tenant Themes
        /// </summary>
        GetTenantThemingOptions,
        /// <summary>
        /// Apply a Theme to the target web
        /// </summary>
        ApplyTheme,
        /// <summary>
        /// Update a Tenant Theme
        /// </summary>
        UpdateTenantTheme,
    }
#endif 
}
