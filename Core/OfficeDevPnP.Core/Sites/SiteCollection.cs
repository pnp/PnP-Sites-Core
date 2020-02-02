#if !ONPREMISES
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Async;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
#if !NETSTANDARD2_0
using System.Text.Encodings.Web;
#endif
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
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static ClientContext Create(ClientContext clientContext, CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation, int delayAfterCreation = 0, bool noWait = false)
        {
            var context = CreateAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation, noWait: noWait).GetAwaiter().GetResult();
            return context;
        }

        /// <summary>
        /// Creates a new Team Site Collection with no group and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static ClientContext Create(ClientContext clientContext, TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation, int delayAfterCreation = 0, bool noWait = false)
        {
            var context = CreateAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation, noWait: noWait).GetAwaiter().GetResult();
            return context;
        }

        /// <summary>
        /// Creates a new Team Site Collection and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static ClientContext Create(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation, int delayAfterCreation = 0, bool noWait = false)
        {
            var context = CreateAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation, noWait: noWait).GetAwaiter().GetResult();
            return context;
        }

        /// <summary>
        /// Creates a new Communication Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>        
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation, int delayAfterCreation = 0, bool noWait = false)
        {
            Dictionary<string, object> payload = GetRequestPayload(siteCollectionCreationInformation);

            var siteDesignId = GetSiteDesignId(siteCollectionCreationInformation);
            if (siteDesignId != Guid.Empty)
            {
                payload.Add("SiteDesignId", siteDesignId);
            }
            payload.Add("HubSiteId", siteCollectionCreationInformation.HubSiteId);

            return await CreateAsync(clientContext, siteCollectionCreationInformation.Owner, payload, delayAfterCreation, noWait: noWait);
        }

        /// <summary>
        /// Creates a new Team Site Collection with no group
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation, int delayAfterCreation = 0, bool noWait = false)
        {
            Dictionary<string, object> payload = GetRequestPayload(siteCollectionCreationInformation);
            return await CreateAsync(clientContext, siteCollectionCreationInformation.Owner, payload, delayAfterCreation, noWait: noWait);
        }

        /// <summary>
        /// Creates a new Modern Team Site Collection (so with an Office 365 group connected)
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0,
            int maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            int retryDelay = 1000 * 10, // Wait time default to 10sec,
            bool noWait = false
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
                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(clientContext);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = string.Format("{0}/_api/GroupSiteManager/CreateGroupEx", clientContext.Web.Url);

                    Dictionary<string, object> payload = new Dictionary<string, object>();
                    payload.Add("displayName", siteCollectionCreationInformation.DisplayName);
                    payload.Add("alias", siteCollectionCreationInformation.Alias);
                    payload.Add("isPublic", siteCollectionCreationInformation.IsPublic);

                    var optionalParams = new Dictionary<string, object>();
                    optionalParams.Add("Description", siteCollectionCreationInformation.Description ?? "");
                    optionalParams.Add("Classification", siteCollectionCreationInformation.Classification ?? "");
                    var creationOptionsValues = new List<string>();
                    if (siteCollectionCreationInformation.SiteDesignId.HasValue)
                    {
                        creationOptionsValues.Add($"implicit_formula_292aa8a00786498a87a5ca52d9f4214a_{siteCollectionCreationInformation.SiteDesignId.Value.ToString("D").ToLower()}");
                    }
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
                    if (MediaTypeHeaderValue.TryParse("application/json;odata.metadata=none;charset=utf-8", out MediaTypeHeaderValue sharePointJsonMediaType))
                    {
                        requestBody.Headers.ContentType = sharePointJsonMediaType;
                    }
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
                        else
                        {
                            if (!noWait)
                            {
                                // Let's wait for the async provisioning of features, site scripts and content types to be done before we allow API's to further update the created site
                                WaitForProvisioningIsComplete(responseContext.Web);
                            }
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
        /// Create a modern site without a group (so communication site and modern team sites without group STS#3)
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="owner">Owner for the created site (needed when using app-only)</param>
        /// <param name="payload">Body of the request</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
        private static async Task<ClientContext> CreateAsync(ClientContext clientContext, string owner, Dictionary<string, object> payload,
            int delayAfterCreation = 0,
            int maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            int retryDelay = 1000 * 10, // Wait time default to 10sec
            bool noWait = false
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
                if (string.IsNullOrEmpty(accessToken))
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
                    if (MediaTypeHeaderValue.TryParse("application/json;odata.metadata=none;charset=utf-8", out MediaTypeHeaderValue sharePointJsonMediaType))
                    {
                        requestBody.Headers.ContentType = sharePointJsonMediaType;
                    }
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
                        else
                        {
                            if (!noWait)
                            {
                                // Let's wait for the async provisioning of features, site scripts and content types to be done before we allow API's to further update the created site
                                WaitForProvisioningIsComplete(responseContext.Web);
                            }
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

        private static void WaitForProvisioningIsComplete(Web web, int maxRetryCount = 80, int retryDelay = 1000 * 15)
        {
            bool isProvisioningComplete = true;
            try
            {
                // Load property
                try
                {
                    web.Context.Load(web, p => p.IsProvisioningComplete);
                    web.Context.ExecuteQueryRetry();
                    isProvisioningComplete = web.IsProvisioningComplete;

                    if (isProvisioningComplete)
                    {
                        // Things went really smooth :-)
                        return;
                    }
                }
                catch (Exception ex)
                {
                    // Catch this...sometimes there's that "sharepoint push feature has not been ..." error
                }

                // Let's start polling for completion. We'll wait maximum 20 minutes for completion.               
                var retryAttempt = 1;
                do
                {
                    if (retryAttempt > 1)
                    {
                        System.Threading.Thread.Sleep(retryDelay);
                    }

                    web.Context.Load(web, p => p.IsProvisioningComplete);
                    web.Context.ExecuteQueryRetry();
                    isProvisioningComplete = web.IsProvisioningComplete;

                    retryAttempt++;

                    // If we already waited more than 90 secs
                    if (retryAttempt * retryDelay > 90000)
                    {
                        var unlockUrl = UrlUtility.Combine(web.Context.Url,
                            "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ValidatePendingWebTemplateExtension");

                        var clientContext = web.Context as ClientContext;

                        HttpHelper.MakePostRequest(unlockUrl, spContext: clientContext);
                    }
                }
                while (!isProvisioningComplete && retryAttempt <= maxRetryCount);
            }
            catch (Exception)
            {
                // Eat the exception for now as not all tenants already have this feature
                // TODO: remove try/catch once IsProvisioningComplete is globally deployed
                isProvisioningComplete = true;
            }

            if (!isProvisioningComplete)
            {
                // Bummer, sites seems to be still not ready...log a warning but let's not fail
                Log.Warning(Constants.LOGGING_SOURCE, string.Format(CoreResources.SiteCollection_WaitForIsProvisioningComplete, maxRetryCount * retryDelay));
                //throw new Exception($"Server side provisioning of this web did not finish after waiting for {maxRetryCount * retryDelay} milliseconds.");
            }
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
                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(clientContext);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = string.Format("{0}/_api/GroupSiteManager/CreateGroupForSite", clientContext.Web.Url);

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
                    if (MediaTypeHeaderValue.TryParse("application/json;odata.metadata=none;charset=utf-8", out MediaTypeHeaderValue sharePointJsonMediaType))
                    {
                        requestBody.Headers.ContentType = sharePointJsonMediaType;
                    }
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

                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string requestUrl = string.Format("{0}/_api/SP.Directory.DirectorySession/Group(alias='{1}')", context.Web.Url, alias);
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
        [Obsolete("Use GetGroupInfoAsync instead of GetGroupInfo")]
        public static async Task<Dictionary<string, string>> GetGroupInfo(ClientContext context, string alias)
        {
            return await GetGroupInfoAsync(context, alias);
        }

        /// <summary>
        /// Checks if a given alias is already in use or not
        /// </summary>
        /// <param name="context">Context to operate against</param>
        /// <param name="alias">Alias to check</param>
        /// <returns>True if in use, false otherwise</returns>
        public static async Task<Dictionary<string, string>> GetGroupInfoAsync(ClientContext context, string alias)
        {
            await new SynchronizationContextRemover();

            Dictionary<string, string> siteInfo = new Dictionary<string, string>();

            var accessToken = context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string requestUrl = string.Format("{0}/_api/SP.Directory.DirectorySession/Group(alias='{1}')", context.Web.Url, alias);
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

        [Obsolete("Use SetGroupImageAsync instead of SetGroupImage")]
        public static async Task<bool> SetGroupImage(ClientContext context, byte[] file, string mimeType)
        {
            return await SetGroupImageAsync(context, file, mimeType);
        }

        /// <summary>
        /// Sets the image for an Office 365 group
        /// </summary>
        /// <param name="context">Context to operate on</param>
        /// <param name="file">Byte array containing the group image</param>
        /// <param name="mimeType">Image mime type</param>
        /// <returns>true if succeeded</returns>
        public static async Task<bool> SetGroupImageAsync(ClientContext context, byte[] file, string mimeType)
        {
            var accessToken = context.GetAccessToken();
            var returnValue = false;
            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(w => w.Url);

                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (string.IsNullOrEmpty(accessToken))
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

                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string requestUrl = string.Format("{0}/_api/GroupSiteManager/GetValidSiteUrlFromAlias?alias='{1}'", context.Web.Url, alias);
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

        /// <summary>
        /// Checks if the Teamify prompt/banner is displayed in the O365 group connected sites.        
        /// </summary>
        /// <param name="context">ClientContext of the site to operate against</param>
        /// <returns></returns>
        public static async Task<bool> IsTeamifyPromptHiddenAsync(ClientContext context)
        {
            bool responseString = false;

            context.Site.EnsureProperties(s => s.GroupId, s => s.Url);

            if (context.Site.GroupId == Guid.Empty)
            {
                throw new Exception("Teamify prompts are only displayed in O365 group connected sites.");
            }
            else
            {
                var result = await context.Web.ExecuteGet($"/_api/groupsitemanager/IsTeamifyPromptHidden?siteUrl='{context.Site.Url}'");

                var teamifyPromptHidden = JObject.Parse(result);

                responseString = Convert.ToBoolean(teamifyPromptHidden["value"]);

                return await Task.Run(() => responseString);
            }
        }

        /// <summary>
        /// Hide the teamify prompt/banner displayed in O365 group connected sites
        /// </summary>
        /// <param name="context">ClientContext of the site to operate against</param>
        /// <returns></returns>
        public static async Task<bool> HideTeamifyPromptAsync(ClientContext context)
        {
            bool responseString = false;

            context.Site.EnsureProperties(s => s.GroupId, s => s.Url);

            if (context.Site.GroupId == Guid.Empty)
            {
                throw new Exception("Teamify prompts can only be hidden in O365 group connected sites.");
            }
            else
            {
                var result = await context.Web.ExecutePost("/_api/groupsitemanager/HideTeamifyPrompt", $@" {{ ""siteUrl"": ""{context.Site.Url}"" }}");

                var teamifyPromptHidden = JObject.Parse(result);

                responseString = Convert.ToBoolean(teamifyPromptHidden["odata.null"]);

                return await Task.Run(() => responseString);
            }
        }

        /// <summary>
        /// Turns a team site into a communication site
        /// </summary>
        /// <param name="context">ClientContext of the team site to update to a communication site</param>
        /// <returns></returns>
        public static async Task EnableCommunicationSite(ClientContext context)
        {
            await EnableCommunicationSite(context, Guid.Parse("96c933ac-3698-44c7-9f4a-5fd17d71af9e"));
        }

        /// <summary>
        /// Turns a team site into a communication site
        /// </summary>
        /// <param name="context">ClientContext of the team site to update to a communication site</param>
        /// <param name="designPackageId">Design package id to be applied, 96c933ac-3698-44c7-9f4a-5fd17d71af9e (Topic = default), 6142d2a0-63a5-4ba0-aede-d9fefca2c767 (Showcase) or f6cc5403-0d63-442e-96c0-285923709ffc (Blank)</param>
        /// <returns></returns>
        public static async Task EnableCommunicationSite(ClientContext context, Guid designPackageId)
        {

            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            context.Web.EnsureProperty(p => p.Url);

            if (designPackageId == Guid.Empty)
            {
                throw new Exception("Please specify a valid designPackageId");
            }

            if (designPackageId != Guid.Parse("96c933ac-3698-44c7-9f4a-5fd17d71af9e") &&  // Topic
                designPackageId != Guid.Parse("6142d2a0-63a5-4ba0-aede-d9fefca2c767") &&  // Showcase
                designPackageId != Guid.Parse("f6cc5403-0d63-442e-96c0-285923709ffc"))    // Blank
            {
                throw new Exception("Invalid designPackageId specified. Use 96c933ac-3698-44c7-9f4a-5fd17d71af9e (Topic = default), 6142d2a0-63a5-4ba0-aede-d9fefca2c767 (Showcase) or f6cc5403-0d63-442e-96c0-285923709ffc (Blank)");
            }

            await context.Web.ExecutePost("/_api/sitepages/communicationsite/enable", $@" {{ ""designPackageId"": ""{designPackageId.ToString()}"" }}");
        }

    }
}
#endif