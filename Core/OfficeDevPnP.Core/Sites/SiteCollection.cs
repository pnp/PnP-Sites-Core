#if !ONPREMISES
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Sites
{

    /// <summary>
    /// This class can be used to create modern site collections
    /// </summary>
    public static class SiteCollection
    {
        /// <summary>
        /// Creates a new Communication Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            ClientContext responseContext = null;

            var accessToken = clientContext.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                clientContext.Web.EnsureProperty(w => w.Url);
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.Credentials = clientContext.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(clientContext.Web.Url), (clientContext.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(clientContext.Web.Url)));
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = String.Format("{0}/_api/sitepages/communicationsite/create", clientContext.Web.Url);

                    var siteDesignId = GetSiteDesignId(siteCollectionCreationInformation);

                    Dictionary<string, object> payload = new Dictionary<string, object>();
                    payload.Add("__metadata", new { type = "SP.Publishing.CommunicationSiteCreationRequest" });
                    payload.Add("Title", siteCollectionCreationInformation.Title);
                    payload.Add("Url", siteCollectionCreationInformation.Url);
                    payload.Add("AllowFileSharingForGuestUsers", siteCollectionCreationInformation.AllowFileSharingForGuestUsers);
                    if (siteDesignId != Guid.Empty)
                    {
                        payload.Add("SiteDesignId", siteDesignId);
                    }
                    payload.Add("Classification", siteCollectionCreationInformation.Classification == null ? "" : siteCollectionCreationInformation.Classification);
                    payload.Add("Description", siteCollectionCreationInformation.Description == null ? "" : siteCollectionCreationInformation.Description);
                    payload.Add("WebTemplateExtensionId", Guid.Empty);

                    var body = new { request = payload };

                    // Serialize request object to JSON
                    var jsonBody = JsonConvert.SerializeObject(body);
                    var requestBody = new StringContent(jsonBody);

                    // Build Http request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Content = requestBody;
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    MediaTypeHeaderValue sharePointJsonMediaType = null;
                    MediaTypeHeaderValue.TryParse("application/json;odata=verbose;charset=utf-8", out sharePointJsonMediaType);
                    requestBody.Headers.ContentType = sharePointJsonMediaType;

                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    requestBody.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigest());

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
                                if (Convert.ToInt32(responseJson["d"]["Create"]["SiteStatus"]) == 2)
                                {
                                    responseContext = clientContext.Clone(responseJson["d"]["Create"]["SiteUrl"].ToString());
                                }
                                else
                                {
                                    throw new Exception(responseString);
                                }
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
                return await Task.Run(() => responseContext);
            }
        }

        /// <summary>
        /// Creates a new Modern Team Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            ClientContext responseContext = null;

            var accessToken = clientContext.GetAccessToken();

            if (!string.IsNullOrEmpty(accessToken))
            {
                throw new Exception("App-Only is currently not supported.");
            }
            using (var handler = new HttpClientHandler())
            {
                clientContext.Web.EnsureProperty(w => w.Url);
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.Credentials = clientContext.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(clientContext.Web.Url), (clientContext.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(clientContext.Web.Url)));
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = String.Format("{0}/_api/GroupSiteManager/CreateGroupEx", clientContext.Web.Url);

                    Dictionary<string, object> payload = new Dictionary<string, object>();
                    payload.Add("displayName", siteCollectionCreationInformation.DisplayName);
                    payload.Add("alias", siteCollectionCreationInformation.Alias);
                    payload.Add("isPublic", siteCollectionCreationInformation.IsPublic);

                    var optionalParams = new Dictionary<string, object>();
                    optionalParams.Add("Description", siteCollectionCreationInformation.Description != null ? siteCollectionCreationInformation.Description : "");
                    optionalParams.Add("CreationOptions", new { results = new object[0], Classification = siteCollectionCreationInformation.Classification != null ? siteCollectionCreationInformation.Classification : "" });

                    payload.Add("optionalParams", optionalParams);

                    var body = payload;

                    // Serialize request object to JSON
                    var jsonBody = JsonConvert.SerializeObject(body);
                    var requestBody = new StringContent(jsonBody);

                    // Build Http request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Content = requestBody;
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    MediaTypeHeaderValue sharePointJsonMediaType = null;
                    MediaTypeHeaderValue.TryParse("application/json;odata=verbose;charset=utf-8", out sharePointJsonMediaType);
                    requestBody.Headers.ContentType = sharePointJsonMediaType;

                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    requestBody.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigest());

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        var responseString = await response.Content.ReadAsStringAsync();
                        var responseJson = JObject.Parse(responseString);
                        if (Convert.ToInt32(responseJson["d"]["CreateGroupEx"]["SiteStatus"]) == 2)
                        {
                            responseContext = clientContext.Clone(responseJson["d"]["CreateGroupEx"]["SiteUrl"].ToString());
                        }
                        else
                        {
                            throw new Exception(responseString);
                        }
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
                return await Task.Run(() => responseContext);
            }
        }

        private static Guid GetSiteDesignId(CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            if (siteCollectionCreationInformation.SiteDesignId != Guid.Empty)
            {
                return siteCollectionCreationInformation.SiteDesignId;
            }
            else
            {
                switch (siteCollectionCreationInformation.SiteDesign)
                {
                    case CommunicationSiteDesign.Topic:
                        {
                            return Guid.Empty;
                        }
                    case CommunicationSiteDesign.Showcase:
                        {
                            return Guid.Parse("6142d2a0-63a5-4ba0-aede-d9fefca2c767");
                        }
                    case CommunicationSiteDesign.Blank:
                        {
                            return Guid.Parse("f6cc5403-0d63-442e-96c0-285923709ffc");
                        }
                }
            }

            return Guid.Empty;
        }


        private static async Task<string> GetValidSiteUrlFromAliasAsync(ClientContext context, string alias)
        {
            string responseString = null;

            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string requestUrl = String.Format("{0}/_api/GroupSiteManager/GetValidSiteUrlFromAlias?alias='{1}'", context.Web.Url, alias);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("accept", "application/json;odata.metadata=minimal");
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Add("odata-version", "4.0");

                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    // Perform actual GET request
                    HttpResponseMessage response = await httpClient.SendAsync(request);

                    if (response.IsSuccessStatusCode)
                    {
                        // If value empty, URL is taken
                        responseString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
                return await Task.Run(() => responseString);
            }
        }
    }
}
#endif