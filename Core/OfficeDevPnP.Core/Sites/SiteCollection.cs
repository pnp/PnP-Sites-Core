#if !ONPREMISES
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Async;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.Core.Sites
{

    /// <summary>
    /// This class can be used to create modern site collections
    /// </summary>
    public static class SiteCollection
    {

        /// <summary>
        /// Creates a new Communication Site Collection and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static ClientContext Create(ClientContext clientContext, CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation, Int32 delayAfterCreation = 0)
        {
            var context = CreateAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation).GetAwaiter().GetResult();
            return context;
        }

        /// <summary>
        /// Creates a new Team Site Collection with no group and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static ClientContext Create(ClientContext clientContext, TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation, Int32 delayAfterCreation = 0)
        {
            var context = CreateAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation).GetAwaiter().GetResult();
            return context;
        }

        /// <summary>
        /// Creates a new Team Site Collection and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static ClientContext Create(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation, Int32 delayAfterCreation = 0)
        {
            var context = CreateAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation).GetAwaiter().GetResult();
            return context;
        }

        /// <summary>
        /// Creates a new Communication Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation, Int32 delayAfterCreation = 0)
        {
            Dictionary<string, object> payload = GetRequestPayload(siteCollectionCreationInformation);

            var siteDesignId = GetSiteDesignId(siteCollectionCreationInformation);
            if (siteDesignId != Guid.Empty)
            {
                payload.Add("SiteDesignId", siteDesignId);
            }
            payload.Add("HubSiteId", siteCollectionCreationInformation.HubSiteId);

            return await CreateAsync(clientContext, siteCollectionCreationInformation.Owner, payload, delayAfterCreation);
        }

        /// <summary>
        /// Creates a new Team Site Collection with no group
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation, Int32 delayAfterCreation = 0)
        {
            Dictionary<string, object> payload = GetRequestPayload(siteCollectionCreationInformation);
            return await CreateAsync(clientContext, siteCollectionCreationInformation.Owner, payload, delayAfterCreation);
        }

        private static Dictionary<string, object> GetRequestPayload(SiteCreationInformation siteCollectionCreationInformation)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>
            {
                { "Title", siteCollectionCreationInformation.Title },
                { "Lcid", siteCollectionCreationInformation.Lcid },
                { "ShareByEmailEnabled", siteCollectionCreationInformation.ShareByEmailEnabled },
                { "Url", siteCollectionCreationInformation.Url },
                { "Classification", siteCollectionCreationInformation.Classification ?? "" },
                { "Description", siteCollectionCreationInformation.Description ?? "" },
                { "WebTemplate", siteCollectionCreationInformation.WebTemplate },
                { "WebTemplateExtensionId", Guid.Empty },
                { "Owner", siteCollectionCreationInformation.Owner }
            };
            return payload;
        }

        private static async Task<ClientContext> CreateAsync(ClientContext clientContext, string owner, Dictionary<string, object> payload, 
            Int32 delayAfterCreation = 0,
            Int32 maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            Int32 retryDelay = 1000 * 10 // Wait time default to 10sec
            )
        {
            await new SynchronizationContextRemover();

            ClientContext responseContext = null;

            if (clientContext.IsAppOnly() && string.IsNullOrEmpty(owner))
            {
                throw new Exception("You need to set the owner in App-only context");
            }

            var accessToken = clientContext.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                clientContext.Web.EnsureProperty(w => w.Url);
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(clientContext);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = $"{clientContext.Web.Url}/_api/SPSiteManager/Create";

                    var body = new { request = payload };

                    // Serialize request object to JSON
                    var jsonBody = JsonConvert.SerializeObject(body);
                    var requestBody = new StringContent(jsonBody);

                    // Build Http request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Content = requestBody;
                    request.Headers.Add("accept", "application/json;odata.metadata=none");
                    request.Headers.Add("odata-version", "4.0");
                    MediaTypeHeaderValue sharePointJsonMediaType = null;
                    MediaTypeHeaderValue.TryParse("application/json;odata.metadata=none;charset=utf-8", out sharePointJsonMediaType);
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
#if !NETSTANDARD2_0
                                if (Convert.ToInt32(responseJson["SiteStatus"]) == 2)
#else
                                if(responseJson["SiteStatus"].Value<int>() == 2)
#endif
                                {
                                    responseContext = clientContext.Clone(responseJson["SiteUrl"].ToString());
                                }
                                else
                                {
                                    /*
                                     * BEGIN : Changes to address the SiteStatus=Provisioning scenario
                                     */
                                    if (Convert.ToInt32(responseJson["SiteStatus"]) == 1)
                                    {
                                        var spOperationsMaxRetryCount = maxRetryCount;
                                        var spOperationsRetryWait = retryDelay;
                                        var siteCreated = false;
                                        var siteUrl = string.Empty;
                                        var retryAttempt = 1;

                                        do
                                        {
                                            if (retryAttempt > 1)
                                            {
                                                System.Threading.Thread.Sleep(retryAttempt * spOperationsRetryWait);
                                            }

                                            try
                                            {
                                                var urlToCheck = HttpUtility.UrlEncode(payload["Url"].ToString());

                                                var siteStatusRequestUrl = $"{clientContext.Web.Url}/_api/SPSiteManager/status?url='{urlToCheck}'";

                                                var siteStatusRequest = new HttpRequestMessage(HttpMethod.Get, siteStatusRequestUrl);
                                                siteStatusRequest.Headers.Add("accept", "application/json;odata=verbose");

                                                if (!string.IsNullOrEmpty(accessToken))
                                                {
                                                    siteStatusRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                                                }

                                                siteStatusRequest.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigest());

                                                var siteStatusResponse = await httpClient.SendAsync(siteStatusRequest, new System.Threading.CancellationToken());
                                                var siteStatusResponseString = await siteStatusResponse.Content.ReadAsStringAsync();

                                                var siteStatusResponseJson = JObject.Parse(siteStatusResponseString);

                                                if (siteStatusResponse.IsSuccessStatusCode)
                                                {
                                                    var siteStatus = Convert.ToInt32(siteStatusResponseJson["d"]["Status"]["SiteStatus"].ToString());
                                                    if (siteStatus == 2)
                                                    {
                                                        siteCreated = true;
                                                        siteUrl = siteStatusResponseJson["d"]["Status"]["SiteUrl"].ToString();
                                                    }
                                                }
                                            }
                                            catch (Exception)
                                            {
                                                // Just skip it and retry after a delay
                                            }

                                            retryAttempt++;
                                        }
                                        while (!siteCreated && retryAttempt <= spOperationsMaxRetryCount);

                                        if (siteCreated)
                                        {
                                            responseContext = clientContext.Clone(siteUrl);
                                        }
                                        else
                                        {
                                            throw new Exception($"OfficeDevPnP.Core.Sites.SiteCollection.CreateAsync: Could not create {payload["WebTemplate"].ToString()} site.");
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception(responseString);
                                    }
                                    /*
                                     * END : Changes to address the SiteStatus=Provisioning scenario
                                     */
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }

                        // If there is a delay, let's wait
                        if (delayAfterCreation > 0)
                        {
                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(delayAfterCreation));
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
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation, 
            Int32 delayAfterCreation = 0, 
            Int32 maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            Int32 retryDelay = 1000 * 10 // Wait time default to 10sec
            )
        {
            if (siteCollectionCreationInformation.Alias.Contains(" "))
            {
                throw new ArgumentException("Alias cannot contain spaces", "Alias");
            }

            await new SynchronizationContextRemover();

            ClientContext responseContext = null;

            if (clientContext.IsAppOnly())
            {
                throw new Exception("App-Only is currently not supported.");
            }

            var accessToken = clientContext.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                clientContext.Web.EnsureProperty(w => w.Url);
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(clientContext);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = String.Format("{0}/_api/GroupSiteManager/CreateGroupEx", clientContext.Web.Url);

                    Dictionary<string, object> payload = new Dictionary<string, object>();
                    payload.Add("displayName", siteCollectionCreationInformation.DisplayName);
                    payload.Add("alias", siteCollectionCreationInformation.Alias);
                    payload.Add("isPublic", siteCollectionCreationInformation.IsPublic);

                    var optionalParams = new Dictionary<string, object>();
                    optionalParams.Add("Description", siteCollectionCreationInformation.Description ?? "");
                    optionalParams.Add("Classification", siteCollectionCreationInformation.Classification ?? "");
                    var creationOptionsValues = new List<string>();
                    if (siteCollectionCreationInformation.Lcid != 0)
                    {
                        creationOptionsValues.Add($"SPSiteLanguage:{siteCollectionCreationInformation.Lcid}");
                    }
                    creationOptionsValues.Add($"HubSiteId:{siteCollectionCreationInformation.HubSiteId}");
                    optionalParams.Add("CreationOptions", creationOptionsValues);

                    if (siteCollectionCreationInformation.Owners != null && siteCollectionCreationInformation.Owners.Length > 0)
                    {
                        optionalParams.Add("Owners", siteCollectionCreationInformation.Owners);
                    }
                    payload.Add("optionalParams", optionalParams);

                    var body = payload;

                    // Serialize request object to JSON
                    var jsonBody = JsonConvert.SerializeObject(body);
                    var requestBody = new StringContent(jsonBody);

                    // Build Http request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Content = requestBody;
                    request.Headers.Add("accept", "application/json;odata.metadata=none");
                    request.Headers.Add("odata-version", "4.0");
                    MediaTypeHeaderValue sharePointJsonMediaType = null;
                    MediaTypeHeaderValue.TryParse("application/json;odata.metadata=none;charset=utf-8", out sharePointJsonMediaType);
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
#if !NETSTANDARD2_0
                        if (Convert.ToInt32(responseJson["SiteStatus"]) == 2)
#else
                        if (responseJson["SiteStatus"].Value<int>() == 2)
#endif
                        {
                            responseContext = clientContext.Clone(responseJson["SiteUrl"].ToString());
                        }
                        else
                        {
                            /*
                             * BEGIN : Changes to address the SiteStatus=Provisioning scenario
                             */
                            if (Convert.ToInt32(responseJson["SiteStatus"]) == 1 && string.IsNullOrWhiteSpace(Convert.ToString(responseJson["ErrorMessage"])))
                            {
                                var spOperationsMaxRetryCount = maxRetryCount; 
                                var spOperationsRetryWait = retryDelay; 
                                var siteCreated = false;
                                var siteUrl = string.Empty;
                                var retryAttempt = 1;

                                do
                                {
                                    if (retryAttempt > 1)
                                    {
                                        System.Threading.Thread.Sleep(retryAttempt * spOperationsRetryWait);
                                    }

                                    try
                                    {
                                        var groupId = responseJson["GroupId"].ToString();
                                        var siteStatusRequestUrl = $"{clientContext.Web.Url}/_api/groupsitemanager/GetSiteStatus('{groupId}')";

                                        var siteStatusRequest = new HttpRequestMessage(HttpMethod.Get, siteStatusRequestUrl);
                                        siteStatusRequest.Headers.Add("accept", "application/json;odata=verbose");

                                        if (!string.IsNullOrEmpty(accessToken))
                                        {
                                            siteStatusRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                                        }

                                        siteStatusRequest.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigest());

                                        var siteStatusResponse = await httpClient.SendAsync(siteStatusRequest, new System.Threading.CancellationToken());
                                        var siteStatusResponseString = await siteStatusResponse.Content.ReadAsStringAsync();

                                        var siteStatusResponseJson = JObject.Parse(siteStatusResponseString);

                                        if (siteStatusResponse.IsSuccessStatusCode)
                                        {
                                            var siteStatus = Convert.ToInt32(siteStatusResponseJson["d"]["GetSiteStatus"]["SiteStatus"].ToString());
                                            if (siteStatus == 2)
                                            {
                                                siteCreated = true;
                                                siteUrl = siteStatusResponseJson["d"]["GetSiteStatus"]["SiteUrl"].ToString();
                                            }
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        // Just skip it and retry after a delay
                                    }

                                    retryAttempt++;
                                }
                                while (!siteCreated && retryAttempt <= spOperationsMaxRetryCount);

                                if (siteCreated)
                                {
                                    responseContext = clientContext.Clone(siteUrl);
                                }
                                else
                                {
                                    throw new Exception("OfficeDevPnP.Core.Sites.SiteCollection.CreateAsync: Could not create team site.");
                                }
                            }
                            else
                            {
                                throw new Exception(responseString);
                            }
                            /*
                             * END : Changes to address the SiteStatus=Provisioning scenario
                             */
                        }

                        // If there is a delay, let's wait
                        if (delayAfterCreation > 0)
                        {
                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(delayAfterCreation));
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
        /// Groupifies a classic team site by creating a group for it and connecting the site with the newly created group
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionGroupifyInformation">information about the site to create</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> GroupifyAsync(ClientContext clientContext, TeamSiteCollectionGroupifyInformation siteCollectionGroupifyInformation)
        {
            if (siteCollectionGroupifyInformation == null)
            {
                throw new ArgumentException("Missing value for siteCollectionGroupifyInformation", "sitecollectionGroupifyInformation");
            }

            if (!string.IsNullOrEmpty(siteCollectionGroupifyInformation.Alias) && siteCollectionGroupifyInformation.Alias.Contains(" "))
            {
                throw new ArgumentException("Alias cannot contain spaces", "Alias");
            }

            if (string.IsNullOrEmpty(siteCollectionGroupifyInformation.DisplayName))
            {
                throw new ArgumentException("DisplayName is required", "DisplayName");
            }

            await new SynchronizationContextRemover();

            ClientContext responseContext = null;

            if (clientContext.IsAppOnly())
            {
                throw new Exception("App-Only is currently not supported.");
            }

            var accessToken = clientContext.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                clientContext.Web.EnsureProperty(w => w.Url);
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(clientContext);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = String.Format("{0}/_api/GroupSiteManager/CreateGroupForSite", clientContext.Web.Url);

                    Dictionary<string, object> payload = new Dictionary<string, object>();
                    payload.Add("displayName", siteCollectionGroupifyInformation.DisplayName);
                    payload.Add("alias", siteCollectionGroupifyInformation.Alias);
                    payload.Add("isPublic", siteCollectionGroupifyInformation.IsPublic);

                    var optionalParams = new Dictionary<string, object>();
                    optionalParams.Add("Description", siteCollectionGroupifyInformation.Description ?? "");
                    optionalParams.Add("Classification", siteCollectionGroupifyInformation.Classification ?? "");
                    // Handle groupify options
                    var creationOptionsValues = new List<string>();
                    if (siteCollectionGroupifyInformation.KeepOldHomePage)
                    {
                        creationOptionsValues.Add("SharePointKeepOldHomepage");
                    }
                    creationOptionsValues.Add($"HubSiteId:{siteCollectionGroupifyInformation.HubSiteId}");
                    optionalParams.Add("CreationOptions", creationOptionsValues);
                    if (siteCollectionGroupifyInformation.Owners != null && siteCollectionGroupifyInformation.Owners.Length > 0)
                    {
                        optionalParams.Add("Owners", siteCollectionGroupifyInformation.Owners);
                    }

                    payload.Add("optionalParams", optionalParams);

                    var body = payload;

                    // Serialize request object to JSON
                    var jsonBody = JsonConvert.SerializeObject(body);
                    var requestBody = new StringContent(jsonBody);

                    // Build Http request
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Content = requestBody;
                    request.Headers.Add("accept", "application/json;odata.metadata=none");
                    request.Headers.Add("odata-version", "4.0");
                    MediaTypeHeaderValue sharePointJsonMediaType = null;
                    MediaTypeHeaderValue.TryParse("application/json;odata.metadata=none;charset=utf-8", out sharePointJsonMediaType);
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

                        // SiteStatus 1 = Provisioning, SiteStatus 2 = Ready
#if !NETSTANDARD2_0
                        if (Convert.ToInt32(responseJson["SiteStatus"]) == 2 || Convert.ToInt32(responseJson["SiteStatus"]) == 1)
#else
                        if (responseJson["SiteStatus"].Value<int>() == 2 || responseJson["SiteStatus"].Value<int>() == 1)
#endif                  
                        {
                            responseContext = clientContext;
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

        /// <summary>
        /// Checks if a given alias is already in use or not
        /// </summary>
        /// <param name="context">Context to operate against</param>
        /// <param name="alias">Alias to check</param>
        /// <returns>True if in use, false otherwise</returns>
        public static async Task<bool> AliasExistsAsync(ClientContext context, string alias)
        {
            await new SynchronizationContextRemover();

            bool aliasExists = true;

            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string requestUrl = String.Format("{0}/_api/SP.Directory.DirectorySession/Group(alias='{1}')", context.Web.Url, alias);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("accept", "application/json;odata.metadata=none");
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Add("odata-version", "4.0");

                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    // Perform actual GET request
                    HttpResponseMessage response = await httpClient.SendAsync(request);

                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        aliasExists = false;
                        // If value empty, URL is taken
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        aliasExists = true;
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
                return await Task.Run(() => aliasExists);
            }
        }

        /// <summary>
        /// Checks if a given alias is already in use or not
        /// </summary>
        /// <param name="context">Context to operate against</param>
        /// <param name="alias">Alias to check</param>
        /// <returns>True if in use, false otherwise</returns>
        public static async Task<Dictionary<string, string>> GetGroupInfo(ClientContext context, string alias)
        {
            await new SynchronizationContextRemover();

            Dictionary<string, string> siteInfo = new Dictionary<string, string>();

            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string requestUrl = String.Format("{0}/_api/SP.Directory.DirectorySession/Group(alias='{1}')", context.Web.Url, alias);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("accept", "application/json;odata.metadata=none");
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Add("odata-version", "4.0");

                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    // Perform actual GET request
                    HttpResponseMessage response = await httpClient.SendAsync(request);

                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        siteInfo = null;
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        var responseString = await response.Content.ReadAsStringAsync();
                        siteInfo = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseString);
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
                return await Task.Run(() => siteInfo);
            }
        }

        public static async Task<bool> SetGroupImage(ClientContext context, byte[] file, string mimeType)
        {
            var accessToken = context.GetAccessToken();
            var returnValue = false;
            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {

                    string requestUrl = $"{context.Web.Url}/_api/groupservice/setgroupimage";

                    var requestDigest = await context.GetRequestDigest();
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    request.Headers.Add("X-RequestDigest", requestDigest);
                    request.Headers.Add("binaryStringRequestBody", "true");
                    request.Content = new ByteArrayContent(file);
                    request.Content.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
                    httpClient.Timeout = new TimeSpan(0, 0, 200);
                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    returnValue = response.IsSuccessStatusCode;
                }
            }
            return await Task.Run(() => returnValue);
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
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string requestUrl = String.Format("{0}/_api/GroupSiteManager/GetValidSiteUrlFromAlias?alias='{1}'", context.Web.Url, alias);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("accept", "application/json;odata.metadata=none");
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

        /// <summary>
        /// Enable Microsoft Teams team in an O365 group connected team site
        /// Will also enable it on a newly Groupified classic site
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public static async Task<string> TeamifySiteAsync(ClientContext context)
        {
            string responseString = null;

            context.Site.EnsureProperties(s => s.GroupId);

            if (context.Web.IsSubSite())
            {
                throw new Exception("You cannot Teamify a subsite");
            }
            else if (context.Site.GroupId == Guid.Empty)
            {
                throw new Exception($"You cannot associate Teams on this site collection. It is only supported for O365 Group connected sites.");
            }
            else
            {
                var result = await context.Web.ExecutePost("/_api/groupsitemanager/EnsureTeamForGroup", string.Empty);

                var teamId = JObject.Parse(result);

                responseString = Convert.ToString(teamId["value"]);

                return await Task.Run(() => responseString);
            }
        }
    }
}
#endif