#if !SP2013 && !SP2016
#if !ONPREMISES
using Microsoft.Graph;
#endif
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Async;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
#if !ONPREMISES
        /// <summary>
        /// Creates a new Communication Site Collection and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
#else
        /// <summary>
        /// Creates a new Communication Site Collection and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
#endif
        public static ClientContext Create(
            ClientContext clientContext,
            CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0
#if !SP2019
            ,
            bool noWait = false
#endif      
            )
        {
            var context = CreateAsync(
                clientContext,
                siteCollectionCreationInformation,
                delayAfterCreation
#if !SP2019
                ,
                noWait: noWait
#endif
                ).GetAwaiter().GetResult();
            return context;
        }

#if !ONPREMISES
        /// <summary>
        /// Creates a new Team Site Collection with no group and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
#else
        /// <summary>
        /// Creates a new Team Site Collection with no group and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
#endif
        public static ClientContext Create(
            ClientContext clientContext,
            TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0
#if !SP2019
            ,
            bool noWait = false
#endif

            )
        {
            var context = CreateAsync(
                clientContext,
                siteCollectionCreationInformation,
                delayAfterCreation
#if !SP2019
                ,
                noWait: noWait
#endif
                ).GetAwaiter().GetResult();
            return context;
        }

#if !ONPREMISES
        /// <summary>
        /// Creates a new Team Site Collection and waits for it to be created
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <param name="graphAccessToken">An optional Access Token for Microsoft Graph to use for creeating the site within an App-Only context</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static ClientContext Create(
            ClientContext clientContext,
            TeamSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0,
            bool noWait = false,
            string graphAccessToken = null)
        {
            var context = CreateAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation, noWait: noWait, graphAccessToken: graphAccessToken).GetAwaiter().GetResult();
            return context;
        }
#endif

#if !ONPREMISES
        /// <summary>
        /// Creates a new Communication Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>        
        /// <returns>ClientContext object for the created site collection</returns>
#else
        /// <summary>
        /// Creates a new Communication Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
#endif
        public static async Task<ClientContext> CreateAsync(
            ClientContext clientContext,
            CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0
#if !SP2019
            ,
            bool noWait = false
#endif
            )
        {
            Dictionary<string, object> payload = GetRequestPayload(siteCollectionCreationInformation);

            var siteDesignId = GetSiteDesignId(siteCollectionCreationInformation);
            if (siteDesignId != Guid.Empty)
            {
                payload.Add("SiteDesignId", siteDesignId);
            }
#if !SP2019
            payload.Add("HubSiteId", siteCollectionCreationInformation.HubSiteId);

            bool sensitivityLabelExists = !string.IsNullOrEmpty(siteCollectionCreationInformation.SensitivityLabel);
            if (sensitivityLabelExists)
            {
                Guid sensitivityLabelId = await GetSensitivityLabelId(clientContext, siteCollectionCreationInformation.SensitivityLabel);
                payload.Add("SensitivityLabel", sensitivityLabelId);
                payload["Classification"] = siteCollectionCreationInformation.SensitivityLabel;
            }
#endif

                return await CreateAsync(clientContext, siteCollectionCreationInformation.Owner, payload, delayAfterCreation
#if !SP2019
                , noWait: noWait
#endif
                );
        }

#if !ONPREMISES
        /// <summary>
        /// Creates a new Team Site Collection with no group
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
#else
        /// <summary>
        /// Creates a new Team Site Collection with no group
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <returns>ClientContext object for the created site collection</returns>
#endif
        public static async Task<ClientContext> CreateAsync(
            ClientContext clientContext,
            TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0
#if !SP2019
            , bool noWait = false
#endif
            )
        {
            Dictionary<string, object> payload = GetRequestPayload(siteCollectionCreationInformation);
            return await CreateAsync(
                clientContext,
                siteCollectionCreationInformation.Owner,
                payload,
                delayAfterCreation
#if !SP2019
                , noWait: noWait
#endif
                );
        }

#if !ONPREMISES
        /// <summary>
        /// Creates a new Modern Team Site Collection (so with an Office 365 group connected)
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <param name="graphAccessToken">An optional Access Token for Microsoft Graph to use for creeating the site within an App-Only context</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateAsync(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0,
            int maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            int retryDelay = 1000 * 10, // Wait time default to 10sec,
            bool noWait = false,
            string graphAccessToken = null
            )
        {
            if (siteCollectionCreationInformation.Alias.Contains(" "))
            {
                throw new ArgumentException("Alias cannot contain spaces", "Alias");
            }

            await new SynchronizationContextRemover();

            ClientContext responseContext = null;

            if (clientContext.IsAppOnly() && string.IsNullOrEmpty(graphAccessToken))
            {
                throw new Exception("App-Only is currently not supported, unless you provide a Microsoft Graph Access Token.");
            }

            // Creating sites through the Microsoft Graph API is preffered. However, if we need to pass in a PreferredDataLocation or a sensitivity label, we need to use the SharePoint API still.
            if (string.IsNullOrEmpty(graphAccessToken) || siteCollectionCreationInformation.PreferredDataLocation.HasValue || !string.IsNullOrEmpty(siteCollectionCreationInformation.SensitivityLabel))
            {
                // Use the regular REST API of SPO to create the modern Team Site
                responseContext = await CreateTeamSiteViaSPOAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation, maxRetryCount, noWait: noWait);
            }
            else
            {
                // Use Microsoft Graph to create the Office 365 Group, and as such the related modern Team Site
                responseContext = await CreateTeamSiteViaGraphAsync(clientContext, siteCollectionCreationInformation, delayAfterCreation, maxRetryCount, noWait: noWait, graphAccessToken: graphAccessToken);
            }

            return responseContext;
        }
#endif

#if !ONPREMISES
        /// <summary>
        /// Private method to create a new Modern Team Site Collection (so with an Office 365 group connected) using SPO REST
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <returns>ClientContext object for the created site collection</returns>
#else
        /// <summary>
        /// Creates a new Modern Team Site Collection (so with an Office 365 group connected)
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <returns>ClientContext object for the created site collection</returns>
#endif
        private static async Task<ClientContext> CreateTeamSiteViaSPOAsync(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0,
            int maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            int retryDelay = 1000 * 10 // Wait time default to 10sec,
#if !SP2019
            ,
            bool noWait = false
#endif
            )
        {
            ClientContext responseContext = null;

            var accessToken = clientContext.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                clientContext.Web.EnsureProperty(w => w.Url);
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(clientContext);
                }

#if !SP2019
                bool sensitivityLabelExists = !string.IsNullOrEmpty(siteCollectionCreationInformation.SensitivityLabel);

                var sensitivityLabelId = Guid.Empty;
                if (sensitivityLabelExists)
                {
                    sensitivityLabelId = await GetSensitivityLabelId(clientContext, siteCollectionCreationInformation.SensitivityLabel);
                }
#endif

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = string.Format("{0}/_api/GroupSiteManager/CreateGroupEx", clientContext.Web.Url);

                    Dictionary<string, object> payload = new Dictionary<string, object>();
                    payload.Add("displayName", siteCollectionCreationInformation.DisplayName);
                    payload.Add("alias", siteCollectionCreationInformation.Alias);
                    payload.Add("isPublic", siteCollectionCreationInformation.IsPublic);

                    var optionalParams = new Dictionary<string, object>();
                    optionalParams.Add("Description", siteCollectionCreationInformation.Description ?? "");
#if !SP2019
                    if (sensitivityLabelExists && sensitivityLabelId != Guid.Empty)
                    {
                        optionalParams.Add("Classification", siteCollectionCreationInformation.SensitivityLabel ?? "");
                    }
                    else
                    {
                        optionalParams.Add("Classification", siteCollectionCreationInformation.Classification ?? "");
                    }
#endif
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

#if !SP2019
                    if (siteCollectionCreationInformation.Owners != null && siteCollectionCreationInformation.Owners.Length > 0)
                    {
                        optionalParams.Add("Owners", siteCollectionCreationInformation.Owners);
                    }
                    payload.Add("optionalParams", optionalParams);
#endif
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
                    else
                    {
                        if (clientContext.Credentials is System.Net.NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }

                    requestBody.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigestAsync());

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
                                        else
                                        {
                                            if (clientContext.Credentials is System.Net.NetworkCredential networkCredential)
                                            {
                                                handler.Credentials = networkCredential;
                                            }
                                        }

                                        siteStatusRequest.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigestAsync());

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
#if !SP2019
                        else
                        {
                            if (!noWait)
                            {
                                // Let's wait for the async provisioning of features, site scripts and content types to be done before we allow API's to further update the created site
                                WaitForProvisioningIsComplete(responseContext.Web);
                            }
                        }
#endif
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

#if !ONPREMISES
        /// <summary>
        /// Creates a new Modern Team Site Collection (so with an Office 365 group connected) using Microsoft Graph
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="siteCollectionCreationInformation">information about the site to create</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <param name="noWait">If specified the site will be created and the process will be finished asynchronously</param>
        /// <param name="graphAccessToken">An optional Access Token for Microsoft Graph to use for creeating the site within an App-Only context</param>
        /// <returns>ClientContext object for the created site collection</returns>
        public static async Task<ClientContext> CreateTeamSiteViaGraphAsync(ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation,
            int delayAfterCreation = 0,
            int maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            int retryDelay = 1000 * 10, // Wait time default to 10sec,
            bool noWait = false,
            string graphAccessToken = null
            )
        {
            ClientContext responseContext = null;

            var group = OfficeDevPnP.Core.Framework.Graph.UnifiedGroupsUtility.CreateUnifiedGroup(
                siteCollectionCreationInformation.DisplayName,
                siteCollectionCreationInformation.Description,
                siteCollectionCreationInformation.Alias,
                graphAccessToken,
                siteCollectionCreationInformation.Owners,
                null, // No members
                isPrivate: !siteCollectionCreationInformation.IsPublic,
                createTeam: false,
                retryCount: maxRetryCount,
                delay: retryDelay);

            if (group != null && !string.IsNullOrEmpty(group.SiteUrl))
            {
                // Try to configure the site/group classification, if any
                if (!string.IsNullOrEmpty(siteCollectionCreationInformation.Classification))
                {
                    await SetTeamSiteClassification(
                        siteCollectionCreationInformation.Classification,
                        group.GroupId,
                        graphAccessToken
                        );
                }

                responseContext = clientContext.Clone(group.SiteUrl);
            }

            return responseContext;
        }
#endif

        private static async Task SetTeamSiteClassification(string classification, string groupId, string graphAccessToken)
        {
            // Patch the created group
            using (var handler = new HttpClientHandler())
            {
                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = $"https://graph.microsoft.com/v1.0/groups/{groupId}";

                    // Serialize request object to JSON
                    var jsonBody = JsonConvert.SerializeObject(new { classification });
                    var requestBody = new StringContent(jsonBody);

                    // Build Http request
                    HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), requestUrl);
                    request.Content = requestBody;
                    if (MediaTypeHeaderValue.TryParse("application/json", out MediaTypeHeaderValue jsonMediaType))
                    {
                        requestBody.Headers.ContentType = jsonMediaType;
                    }
                    if (!string.IsNullOrEmpty(graphAccessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
                    }

                    // Perform actual post operation
                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception("Failed to set Classification for created group");
                    }
                }
            }
        }

#if !ONPREMISES
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
#else
        /// <summary>
        /// Create a modern site without a group (so communication site and modern team sites without group STS#3)
        /// </summary>
        /// <param name="clientContext">ClientContext object of a regular site</param>
        /// <param name="owner">Owner for the created site (needed when using app-only)</param>
        /// <param name="payload">Body of the request</param>
        /// <param name="delayAfterCreation">Defines the number of seconds to wait after creation</param>
        /// <param name="maxRetryCount">Maximum number of retries for a pending site provisioning. Default 12 retries.</param>
        /// <param name="retryDelay">Delay between retries for a pending site provisioning. Default 10 seconds.</param>
        /// <returns>ClientContext object for the created site collection</returns>
#endif
        private static async Task<ClientContext> CreateAsync(
            ClientContext clientContext,
            string owner,
            Dictionary<string, object> payload,
            int delayAfterCreation = 0,
            int maxRetryCount = 12, // Maximum number of retries (12 x 10 sec = 120 sec = 2 mins)
            int retryDelay = 1000 * 10 // Wait time default to 10sec
#if !SP2019
            ,
            bool noWait = false
#endif
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
                else
                {
                    if (clientContext.Credentials is System.Net.NetworkCredential networkCredential)
                    {
                        handler.Credentials = networkCredential;
                    }
                }

#if SP2019
                // TODO: Check Owner property support in SP2019
                /*
                if (clientContext.IsAppOnly()
                    && !string.IsNullOrEmpty(owner))
                {
                    //"The property 'Owner' does not exist on type 'Microsoft.SharePoint.Portal.SPSiteCreationRequest'. Make sure to only use property names that are defined by the type."
                    throw new Exception("The Owner cannot be set in an App-Only context. Cause SharePoint 2019 server side does not support the owner property yet.");
                }
                */
#endif

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
                    else
                    {
                        if (clientContext.Credentials is System.Net.NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }

                    requestBody.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigestAsync());

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

                                                siteStatusRequest.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigestAsync());

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
                                            var errorSb = new System.Text.StringBuilder();
                                            errorSb.AppendLine($"Result:{responseString}");

                                            //var System.Net.Http.HttpResponseMessage
                                            //if(response.Headers["SPRequestGuid"] != null)
                                            //if (response.Headers.AllKeys.Any(k => string.Equals(k, "SPRequestGuid", StringComparison.InvariantCultureIgnoreCase)))
                                            if (response.Headers.Contains("SPRequestGuid"))
                                            {
                                                var values = response.Headers.GetValues("SPRequestGuid");
                                                if (values != null)
                                                {
                                                    var spRequestGuid = values.FirstOrDefault();
                                                    errorSb.AppendLine($"ServerErrorTraceCorrelationId: {spRequestGuid}");
                                                }
                                            }

                                            clientContext.Web.EnsureProperty(w => w.CurrentUser);
                                            clientContext.Web.CurrentUser.EnsureProperty(u => u.LoginName);
                                            errorSb.AppendLine($"CurrentUser / Owner: {clientContext.Web.CurrentUser.LoginName}");
#if SP2019

                                            if (clientContext.Web.CurrentUser.LoginName.Contains("app@sharepoint"))
                                            {
                                                errorSb.AppendLine($"Please check the account used as owner for the site to create."
                                                + $" Consider that the AppPrincipal for App-Only '{clientContext.Web.CurrentUser.LoginName}' cannot be used as site owner."
                                                + " Setting the 'owner' property does not work. It will lead to the following exception: "
                                                + " The property 'Owner' does not exist on type 'Microsoft.SharePoint.Portal.SPSiteCreationRequest'.Make sure to only use property names that are defined by the type.");

                                            }
#endif
                                            Log.Error(Constants.LOGGING_SOURCE, CoreResources.ClientContextExtensions_ExecuteQueryRetryException, errorSb.ToString());

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
#if !SP2019
                        else
                        {
                            if (!noWait)
                            {
                                // Let's wait for the async provisioning of features, site scripts and content types to be done before we allow API's to further update the created site
                                WaitForProvisioningIsComplete(responseContext.Web);
                            }
                        }
#endif
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

#if !SP2019
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

                Log.Debug(Constants.LOGGING_SOURCE, $"Starting to wait for site collection to be created");

                var stopwatch = new Stopwatch();
                stopwatch.Start();

                var retryAttempt = 1;
                do
                {
                    Log.Debug(Constants.LOGGING_SOURCE, $"Elapsed: {stopwatch.Elapsed.ToString(@"mm\:ss\.fff")} | Attempt {retryAttempt}/{maxRetryCount}");

                    if (retryAttempt > 1)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, $"Elapsed: {stopwatch.Elapsed.ToString(@"mm\:ss\.fff")} | Waiting {retryDelay / 1000} seconds");

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

                stopwatch.Stop();
                Log.Debug(Constants.LOGGING_SOURCE, $"Elapsed: {stopwatch.Elapsed.ToString(@"mm\:ss\.fff")} | Finished");
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
#endif

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
#if !SP2019 // SP2019: Microsoft.OData.Core.ODataException: The property 'Owner' does not exist on type 'Microsoft.SharePoint.Portal.SPSiteCreationRequest'. Make sure to only use property names that are defined by the type.
                { "Owner", siteCollectionCreationInformation.Owner }
#endif
            };
            return payload;
        }

#if !ONPREMISES
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
                else
                {
                    if (clientContext.Credentials is System.Net.NetworkCredential networkCredential)
                    {
                        handler.Credentials = networkCredential;
                    }
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
                    else
                    {
                        if (clientContext.Credentials is System.Net.NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }

                    requestBody.Headers.Add("X-RequestDigest", await clientContext.GetRequestDigestAsync());

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
#endif

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

#if !ONPREMISES
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
                    else
                    {
                        if (context.Credentials is System.Net.NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
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

                    var requestDigest = await context.GetRequestDigestAsync();
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

        /// <summary>
        /// Allows validation if the provided <paramref name="alias"/> is valid to be used to create a new site collection
        /// </summary>
        /// <param name="context">SharePoint ClientContext to use to communicate with SharePoint</param>
        /// <param name="alias">The alias to check for availability</param>
        /// <returns>True if the provided alias is available to be used, false if it is not</returns>
        public static async Task<bool> GetIsAliasAvailableAsync(ClientContext context, string alias)
        {
            var proposedUrl = await GetValidSiteUrlFromAliasAsync(context, alias);
            return proposedUrl.EndsWith($"/{alias}", StringComparison.InvariantCultureIgnoreCase);
        }

        /// <summary>
        /// Checks if the provided <paramref name="alias"/> is valid to be used to create a new site collection and will return an alternative available proposal if it is not. Use <see cref="GetIsAliasAvailableAsync"/> instead if you are just intered in knowing whether or not a certain alias is still available to be used.
        /// </summary>
        /// <param name="context">SharePoint ClientContext to use to communicate with SharePoint</param>
        /// <param name="alias">The alias to check for availability</param>
        /// <returns>The full SharePoint URL proposed to be used. If that URL ends with the alias you provided, it means it is still available. If its not available, it will return an alternative proposal to use.</returns>
        public static async Task<string> GetValidSiteUrlFromAliasAsync(ClientContext context, string alias)
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
                        var requestResponse = await response.Content.ReadAsStringAsync();
                        var requestResponseJson = JObject.Parse(requestResponse);

                        responseString = requestResponseJson["value"].ToString();
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
        /// <param name="context">Context to operate against</param>
        /// <returns></returns>
        public static async Task<string> TeamifySiteAsync(ClientContext context)
        {
            string responseString = null;

            context.Site.EnsureProperty(s => s.GroupId);

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
                throw new Exception("Teamify prompts can only be hidden in Microsoft 365 group connected sites.");
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


        /// <summary>
        /// Get sensitivity label id for a given Label
        /// </summary>
        /// <param name="context">Client context</param>
        /// <param name="sensitiveLabelString">Sensitive Label string value</param>
        /// <returns></returns>
        private static async Task<Guid> GetSensitivityLabelId(ClientContext context, string sensitiveLabelString)
        {
            var result = await context.Web.ExecuteGet("/_api/groupsitemanager/GetGroupCreationContext");

            var results = JObject.Parse(result);

            JToken val = results["DataClassificationOptionsNew"]?.Children().FirstOrDefault(jt => (string)jt["Value"] == sensitiveLabelString);

            string sensitivityLabelStringId = Convert.ToString(val?["Key"]);

            Guid sensitivityLabelId = Guid.Empty;

            if (!string.IsNullOrEmpty(sensitivityLabelStringId))
            {
                sensitivityLabelId = Guid.Parse(sensitivityLabelStringId);
            }

            return await Task.Run(() => sensitivityLabelId);
        }


        /// <summary>
        /// Gets group alias information by group Id
        /// </summary>
        /// <param name="context">Context to operate against</param>
        /// <param name="groupId">Id of the group</param>
        /// <returns>True if in use, false otherwise</returns>
        public static async Task<Dictionary<string, object>> GetGroupInfoByGroupIdAsync(ClientContext context, string groupId)
        {
            await new SynchronizationContextRemover();

            Dictionary<string, object> siteInfo = new Dictionary<string, object>();

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
                    string requestUrl = string.Format("{0}/_api/SP.Directory.DirectorySession/Group('{1}')?$select=PrincipalName,Id,DisplayName,Alias,Description,InboxUrl,CalendarUrl,DocumentsUrl,SiteUrl,EditGroupUrl,PictureUrl,PeopleUrl,NotebookUrl,Mail,IsPublic,CreationTime,Classification,teamsResources,yammerResources,allowToAddGuests,isDynamic,assignedLabels", context.Web.Url, groupId);
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
                        siteInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(responseString);
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

        /// <summary>
        /// Deletes a Communication site or a Group-less Modern team site.
        /// </summary>
        /// <param name="context">Context to operate against</param>
        /// <returns></returns>
        public static async Task<bool> DeleteSiteAsync(ClientContext context)
        {
            bool siteDeleted = false;

            var webTemplateId = context.Web.GetBaseTemplateId();

            context.Site.EnsureProperties(s => s.Id, s => s.GroupId, s => s.Url);

            if (webTemplateId == "SITEPAGEPUBLISHING#0" || webTemplateId == "STS#3")
            {
                var result = await context.Web.ExecutePost("/_api/SPSiteManager/delete", $@" {{ ""siteId"": ""{context.Site.Id.ToString()}"" }}");

                var parsedResult = JObject.Parse(result);

                siteDeleted = Convert.ToBoolean(parsedResult["odata.null"]);

                return await Task.Run(() => siteDeleted);
            }
            else if (webTemplateId == "GROUP#0" || context.Site.GroupId != Guid.Empty)
            {
                var result = await context.Web.ExecutePost($"/_api/GroupSiteManager/Delete?siteUrl='{context.Site.Url}'", string.Empty);

                var parsedResult = JObject.Parse(result);

                siteDeleted = Convert.ToBoolean(parsedResult["odata.null"]);

                return await Task.Run(() => siteDeleted);
            }
            else
            {
                throw new Exception("Only deletion of Communication site or Modern team site is supported by this method.");
            }
        }
#endif
    }
}
#endif