using System;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Web;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using System.Configuration;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Utilities.Async;
using System.IdentityModel.Tokens.Jwt;
using System.Collections.Generic;
using OfficeDevPnP.Core.Utilities.Context;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;

#if !ONPREMISES
using OfficeDevPnP.Core.Sites;
#endif

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that deals with cloning client context object, getting access token and validates server version
    /// </summary>
    public static partial class ClientContextExtensions
    {
        private static string userAgentFromConfig = null;
        private static string accessToken = null;
        private static bool hasAuthCookies;

        /// <summary>
        /// Static constructor, only executed once per class load
        /// </summary>
        static ClientContextExtensions()
        {
            ClientContextExtensions.userAgentFromConfig = ConfigurationManager.AppSettings["SharePointPnPUserAgent"];
        }


#if ONPREMISES
        private const string MicrosoftSharePointTeamServicesHeader = "MicrosoftSharePointTeamServices";
#endif

        /// <summary>
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="siteUrl">Site URL to be used for cloned ClientContext</param>
        /// <param name="accessTokens">Dictionary of access tokens for sites URLs</param>
        /// <returns>A ClientContext object created for the passed site URL</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, string siteUrl, Dictionary<String, String> accessTokens = null)
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUrl));
            }

            return clientContext.Clone(new Uri(siteUrl), accessTokens);
        }

#if !ONPREMISES
        /// <summary>
        /// Executes the current set of data retrieval queries and method invocations and retries it if needed using the Task Library.
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"></param>
        public static Task ExecuteQueryRetryAsync(this ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500, string userAgent = null)
        {
            return ExecuteQueryImplementation(clientContext, retryCount, delay, userAgent);
        }

#endif

        /// <summary>
        /// Executes the current set of data retrieval queries and method invocations and retries it if needed.
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"></param>
        public static void ExecuteQueryRetry(this ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500, string userAgent = null)
        {
#if !ONPREMISES            
            Task.Run(() => ExecuteQueryImplementation(clientContext, retryCount, delay, userAgent)).GetAwaiter().GetResult();
#else
            ExecuteQueryImplementation(clientContext, retryCount, delay, userAgent);
#endif
        }

#if !ONPREMISES
        private static async Task ExecuteQueryImplementation(ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500, string userAgent = null)
#else
        private static void ExecuteQueryImplementation(ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500, string userAgent = null)
#endif
        {

#if !ONPREMISES
            await new SynchronizationContextRemover();

            // Set the TLS preference. Needed on some server os's to work when Office 365 removes support for TLS 1.0
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
#endif

            var clientTag = string.Empty;
            if (clientContext is PnPClientContext)
            {
                retryCount = (clientContext as PnPClientContext).RetryCount;
                delay = (clientContext as PnPClientContext).Delay;
                clientTag = (clientContext as PnPClientContext).ClientTag;
            }

            int retryAttempts = 0;
            int backoffInterval = delay;

#if !ONPREMISES
            int retryAfterInterval = 0;
            bool retry = false;
            ClientRequestWrapper wrapper = null;
#endif

            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");

            if (delay <= 0)
                throw new ArgumentException("Provide a delay greater than zero.");

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
                {
                    clientContext.ClientTag = SetClientTag(clientTag);

                    // Make CSOM request more reliable by disabling the return value cache. Given we 
                    // often clone context objects and the default value is
#if !ONPREMISES || SP2016 || SP2019
                    clientContext.DisableReturnValueCache = true;
#endif
                    // Add event handler to "insert" app decoration header to mark the PnP Sites Core library as a known application
                    EventHandler<WebRequestEventArgs> appDecorationHandler = AttachRequestUserAgent(userAgent);

                    clientContext.ExecutingWebRequest += appDecorationHandler;

                    // DO NOT CHANGE THIS TO EXECUTEQUERYRETRY
#if !ONPREMISES
                    if (!retry)
                    {
#if !NETSTANDARD2_0
                        await clientContext.ExecuteQueryAsync();
#else
                        clientContext.ExecuteQuery();
#endif
                    }
                    else
                    {
                        if (wrapper != null && wrapper.Value != null)
                        {
#if !NETSTANDARD2_0
                            await clientContext.RetryQueryAsync(wrapper.Value);
#else
                            clientContext.RetryQuery(wrapper.Value);
#endif
                        }
                    }
#else
                    clientContext.ExecuteQuery();
#endif

                    // Remove the app decoration event handler after the executequery
                    clientContext.ExecutingWebRequest -= appDecorationHandler;

                    return;
                }
                catch (WebException wex)
                {
                    var response = wex.Response as HttpWebResponse;
                    // Check if request was throttled - http status code 429
                    // Check is request failed due to server unavailable - http status code 503
                    if (response != null && (response.StatusCode == (HttpStatusCode)429 || response.StatusCode == (HttpStatusCode)503))
                    {
                        Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ClientContextExtensions_ExecuteQueryRetry, backoffInterval);

#if !ONPREMISES
                        wrapper = (ClientRequestWrapper)wex.Data["ClientRequest"];
                        retry = true;

                        // Determine the retry after value - use the retry-after header when available
                        // Retry-After seems to default to a fixed 120 seconds in most cases, let's revert to our
                        // previous logic
                        //string retryAfterHeader = response.GetResponseHeader("Retry-After");
                        //if (!string.IsNullOrEmpty(retryAfterHeader))
                        //{
                        //    if (!Int32.TryParse(retryAfterHeader, out retryAfterInterval))
                        //    {
                        //        retryAfterInterval = backoffInterval;
                        //    }
                        //}
                        //else
                        //{
                        retryAfterInterval = backoffInterval;
                        //}

                        //Add delay for retry, retry-after header is specified in seconds
                        await Task.Delay(retryAfterInterval);
#else
                        Thread.Sleep(backoffInterval);
#endif

                        //Add to retry count and increase delay.
                        retryAttempts++;
                        backoffInterval = backoffInterval * 2;
                    }
                    else
                    {
                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.ClientContextExtensions_ExecuteQueryRetryException, wex.ToString());
                        throw;
                    }
                }
            }
            throw new MaximumRetryAttemptedException($"Maximum retry attempts {retryCount}, has be attempted.");
        }

        /// <summary>
        /// Attaches either a passed user agent, or one defined in the App.config file, to the WebRequstExecutor UserAgent property.
        /// </summary>
        /// <param name="customUserAgent">a custom user agent to override any defined in App.config</param>
        /// <returns>An EventHandler of WebRequestEventArgs.</returns>
        private static EventHandler<WebRequestEventArgs> AttachRequestUserAgent(string customUserAgent)
        {
            return (s, e) =>
            {
                bool overrideUserAgent = true;
                var existingUserAgent = e.WebRequestExecutor.WebRequest.UserAgent;
                if (!string.IsNullOrEmpty(existingUserAgent) && existingUserAgent.StartsWith("NONISV|SharePointPnP|PnPPS/"))
                {
                    overrideUserAgent = false;
                }
                if (overrideUserAgent)
                {
                    if (string.IsNullOrEmpty(customUserAgent) && !string.IsNullOrEmpty(ClientContextExtensions.userAgentFromConfig))
                    {
                        customUserAgent = userAgentFromConfig;
                    }
                    e.WebRequestExecutor.WebRequest.UserAgent = string.IsNullOrEmpty(customUserAgent) ? $"{PnPCoreUtilities.PnPCoreUserAgent}" : customUserAgent;
                }
            };
        }

        /// <summary>
        /// Sets the client context client tag on outgoing CSOM requests.
        /// </summary>
        /// <param name="clientTag">An optional client tag to set on client context requests.</param>
        /// <returns></returns>
        private static string SetClientTag(string clientTag = "")
        {
            // ClientTag property is limited to 32 chars
            if (string.IsNullOrEmpty(clientTag))
            {
                clientTag = $"{PnPCoreUtilities.PnPCoreVersionTag}:{GetCallingPnPMethod()}";
            }
            if (clientTag.Length > 32)
            {
                clientTag = clientTag.Substring(0, 32);
            }

            return clientTag;
        }


        /// <summary>
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="siteUrl">Site URL to be used for cloned ClientContext</param>
        /// <param name="accessTokens">Dictionary of access tokens for sites URLs</param>
        /// <returns>A ClientContext object created for the passed site URL</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, Uri siteUrl, Dictionary<String, String> accessTokens = null)
        {
            if (siteUrl == null)
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUrl));
            }

            ClientContext clonedClientContext = new ClientContext(siteUrl);
            clonedClientContext.AuthenticationMode = clientContext.AuthenticationMode;
            clonedClientContext.ClientTag = clientContext.ClientTag;
#if !ONPREMISES || SP2016 || SP2019
            clonedClientContext.DisableReturnValueCache = clientContext.DisableReturnValueCache;
#endif


            // In case of using networkcredentials in on premises or SharePointOnlineCredentials in Office 365
            if (clientContext.Credentials != null)
            {
                clonedClientContext.Credentials = clientContext.Credentials;
            }
            else
            {

                // Check if we do have context settings
                var contextSettings = clientContext.GetContextSettings();

                if (contextSettings != null) // We do have more information about this client context, so let's use it to do a more intelligent clone
                {
                    string newSiteUrl = siteUrl.ToString();

                    // A diffent host = different audience ==> new access token is needed
                    if (contextSettings.UsesDifferentAudience(newSiteUrl))
                    {
                        // We need to create a new context using a new authentication manager as the token expiration is handled in there
                        OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();

                        ClientContext newClientContext = null;
                        if (contextSettings.Type == ClientContextType.SharePointACSAppOnly)
                        {
                            newClientContext = authManager.GetAppOnlyAuthenticatedContext(newSiteUrl, TokenHelper.GetRealmFromTargetUrl(new Uri(newSiteUrl)), contextSettings.ClientId, contextSettings.ClientSecret, contextSettings.AcsHostUrl, contextSettings.GlobalEndPointPrefix);
                        }
#if !ONPREMISES && !NETSTANDARD2_0
                        else if (contextSettings.Type == ClientContextType.AzureADCredentials)
                        {
                            newClientContext = authManager.GetAzureADCredentialsContext(newSiteUrl, contextSettings.UserName, contextSettings.Password);
                        }
                        else if (contextSettings.Type == ClientContextType.AzureADCertificate)
                        {
                            newClientContext = authManager.GetAzureADAppOnlyAuthenticatedContext(newSiteUrl, contextSettings.ClientId, contextSettings.Tenant, contextSettings.Certificate, contextSettings.Environment);
                        }
#endif

                        if (newClientContext != null)
                        {
                            //Take over the form digest handling setting
                            newClientContext.FormDigestHandlingEnabled = (clientContext as ClientContext).FormDigestHandlingEnabled;
                            newClientContext.ClientTag = clientContext.ClientTag;
#if !ONPREMISES || SP2016 || SP2019
                            newClientContext.DisableReturnValueCache = clientContext.DisableReturnValueCache;
#endif
                            return newClientContext;
                        }
                        else
                        {
                            throw new Exception($"Cloning for context setting type {contextSettings.Type} was not yet implemented");
                        }
                    }
                    else
                    {
                        // Take over the context settings, this is needed if we later on want to clone this context to a different audience
                        contextSettings.SiteUrl = newSiteUrl;
                        clonedClientContext.AddContextSettings(contextSettings);

                        clonedClientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                        {
                            // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                            // the new delegate method
                            MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                            object[] parametersArray = new object[] { webRequestEventArgs };
                            methodInfo.Invoke(clientContext, parametersArray);
                        };
                    }
                }
                else // Fallback the default cloning logic if there were not context settings available
                {
                    //Take over the form digest handling setting
                    clonedClientContext.FormDigestHandlingEnabled = (clientContext as ClientContext).FormDigestHandlingEnabled;

                    var originalUri = new Uri(clientContext.Url);
                    // If the cloned host is not the same as the original one
                    // and if there is an active PnPProvisioningContext
                    if (originalUri.Host != siteUrl.Host &&
                        PnPProvisioningContext.Current != null)
                    {
                        // Let's apply that specific Access Token
                        clonedClientContext.ExecutingWebRequest += (sender, args) =>
                        {
                            // We get a fresh new Access Token for every request, to avoid using an expired one
                            var accessToken = PnPProvisioningContext.Current.AcquireToken(siteUrl.Authority, null);
                            args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                        };
                    }
                    // Else if the cloned host is not the same as the original one
                    // and if there is a custom Access Token for it in the input arguments
                    else if (originalUri.Host != siteUrl.Host &&
                        accessTokens != null && accessTokens.Count > 0 &&
                        accessTokens.ContainsKey(siteUrl.Authority))
                    {
                        // Let's apply that specific Access Token
                        clonedClientContext.ExecutingWebRequest += (sender, args) =>
                        {
                            args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessTokens[siteUrl.Authority];
                        };
                    }
                    // Else if the cloned host is not the same as the original one
                    // and if the client context is a PnPClientContext with custom access tokens in its property bag
                    else if (originalUri.Host != siteUrl.Host &&
                        accessTokens == null && clientContext is PnPClientContext &&
                        ((PnPClientContext)clientContext).PropertyBag.ContainsKey("AccessTokens") &&
                        ((Dictionary<string, string>)((PnPClientContext)clientContext).PropertyBag["AccessTokens"]).ContainsKey(siteUrl.Authority))
                    {
                        // Let's apply that specific Access Token
                        clonedClientContext.ExecutingWebRequest += (sender, args) =>
                        {
                            args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ((Dictionary<string, string>)((PnPClientContext)clientContext).PropertyBag["AccessTokens"])[siteUrl.Authority];
                        };
                    }
                    else
                    {
                        // In case of app only or SAML
                        clonedClientContext.ExecutingWebRequest += (sender, webRequestEventArgs) =>
                        {
                            // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                            // the new delegate method
                            MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                            object[] parametersArray = new object[] { webRequestEventArgs };
                            methodInfo.Invoke(clientContext, parametersArray);
                        };
                    }
                }
            }

            return clonedClientContext;
        }

        /// <summary>
        /// Returns the number of pending requests
        /// </summary>
        /// <param name="clientContext">Client context to check the pending requests for</param>
        /// <returns>The number of pending requests</returns>
        public static int PendingRequestCount(this ClientRuntimeContext clientContext)
        {
            int count = 0;

            if (clientContext.HasPendingRequest)
            {
                var result = clientContext.PendingRequest.GetType().GetProperty("Actions", BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.NonPublic);
                if (result != null)
                {
                    var propValue = result.GetValue(clientContext.PendingRequest);
                    if (propValue != null)
                    {
                        count = (propValue as System.Collections.Generic.List<ClientAction>).Count;
                    }
                }
            }

            return count;
        }

        /// <summary>
        /// Gets a site collection context for the passed web. This site collection client context uses the same credentials
        /// as the passed client context
        /// </summary>
        /// <param name="clientContext">Client context to take the credentials from</param>
        /// <returns>A site collection client context object for the site collection</returns>
        public static ClientContext GetSiteCollectionContext(this ClientRuntimeContext clientContext)
        {
            Site site = (clientContext as ClientContext).Site;
            if (!site.IsObjectPropertyInstantiated("Url"))
            {
                clientContext.Load(site);
                clientContext.ExecuteQueryRetry();
            }
            return clientContext.Clone(site.Url);
        }

        /// <summary>
        /// Checks if the used ClientContext is app-only
        /// </summary>
        /// <param name="clientContext">The ClientContext to inspect</param>
        /// <returns>True if app-only, false otherwise</returns>
        public static bool IsAppOnly(this ClientRuntimeContext clientContext)
        {
            // Set initial result to false
            var result = false;

            // Try to get an access token from the current context
            var accessToken = clientContext.GetAccessToken();

            // If any
            if (!String.IsNullOrEmpty(accessToken))
            {
                // Try to decode the access token
                var token = new JwtSecurityToken(accessToken);

                // Search for the UPN claim, to see if we have user's delegation
                var upn = token.Claims.FirstOrDefault(claim => claim.Type == "upn")?.Value;
                if (String.IsNullOrEmpty(upn))
                {
                    result = true;
                }
            }
            else if (clientContext.Credentials == null)
            {
                result = true;
            }
            // As a final check, do we have the auth cookies?
            if (clientContext.HasAuthCookies())
            {
                result = false;
            }
            return (result);
        }

        /// <summary>
        /// Gets an access token from a <see cref="ClientContext"/> instance. Only works when using an add-in or app-only authentication flow.
        /// </summary>
        /// <param name="clientContext"><see cref="ClientContext"/> instance to obtain an access token for</param>
        /// <returns>Access token for the given <see cref="ClientContext"/> instance</returns>
        public static string GetAccessToken(this ClientRuntimeContext clientContext)
        {
            string accessToken = null;
            EventHandler<WebRequestEventArgs> handler = (s, e) =>
            {
                string authorization = e.WebRequestExecutor.RequestHeaders["Authorization"];
                if (!string.IsNullOrEmpty(authorization))
                {
                    accessToken = authorization.Replace("Bearer ", string.Empty);
                }
            };
            // Issue a dummy request to get it from the Authorization header
            clientContext.ExecutingWebRequest += handler;
            clientContext.ExecuteQueryRetry();
            clientContext.ExecutingWebRequest -= handler;

            return accessToken;
        }

        /// <summary>
        /// Gets a boolean if the current request contains the FedAuth and rtFa cookies.
        /// </summary>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        private static bool HasAuthCookies(this ClientRuntimeContext clientContext)
        {
            clientContext.ExecutingWebRequest += ClientContext_ExecutingWebRequestCookieCounter;
            clientContext.ExecuteQueryRetry();
            clientContext.ExecutingWebRequest -= ClientContext_ExecutingWebRequestCookieCounter;
            return hasAuthCookies;
        }

        private static void ClientContext_ExecutingWebRequestCookieCounter(object sender, WebRequestEventArgs e)
        {
            var fedAuth = false;
            var rtFa = false;

            if (e.WebRequestExecutor != null && e.WebRequestExecutor.WebRequest != null && e.WebRequestExecutor.WebRequest.CookieContainer != null)
            {
                var cookies = e.WebRequestExecutor.WebRequest.CookieContainer.GetCookies(e.WebRequestExecutor.WebRequest.RequestUri);
                if (cookies.Count > 0)
                {
                    for (var q = 0; q < cookies.Count; q++)
                    {
                        if (cookies[q].Name == "FedAuth")
                        {
                            fedAuth = true;
                        }
                        if (cookies[q].Name == "rtFa")
                        {
                            rtFa = true;
                        }
                    }
                }
            }

            hasAuthCookies = fedAuth && rtFa;
        }

        /// <summary>
        /// Defines a Maximum Retry Attemped Exception
        /// </summary>
        [Serializable]
        public class MaximumRetryAttemptedException : Exception
        {
            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="message"></param>
            public MaximumRetryAttemptedException(string message)
                : base(message)
            {

            }
        }

        /// <summary>
        /// Checks the server library version of the context for a minimally required version
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="minimallyRequiredVersion">provide version to validate</param>
        /// <returns>True if it has minimal required version, false otherwise</returns>
        public static bool HasMinimalServerLibraryVersion(this ClientRuntimeContext clientContext, string minimallyRequiredVersion)
        {
            return HasMinimalServerLibraryVersion(clientContext, new Version(minimallyRequiredVersion));
        }

        /// <summary>
        /// Checks the server library version of the context for a minimally required version
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="minimallyRequiredVersion">provide version to validate</param>
        /// <returns>True if it has minimal required version, false otherwise</returns>
        public static bool HasMinimalServerLibraryVersion(this ClientRuntimeContext clientContext, Version minimallyRequiredVersion)
        {
            bool hasMinimalVersion = false;
#if !ONPREMISES
            try
            {
                clientContext.ExecuteQueryRetry();
                hasMinimalVersion = clientContext.ServerLibraryVersion.CompareTo(minimallyRequiredVersion) >= 0;
            }
            catch (PropertyOrFieldNotInitializedException)
            {
                // swallow the exception.
            }
#else
            try
            {
                Uri urlUri = new Uri(clientContext.Url);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create($"{urlUri.Scheme}://{urlUri.DnsSafeHost}:{urlUri.Port}/_vti_pvt/service.cnf");
                request.UseDefaultCredentials = true;

                var response = request.GetResponse();

                using (var dataStream = response.GetResponseStream())
                {
                    // Open the stream using a StreamReader for easy access.
                    using (System.IO.StreamReader reader = new System.IO.StreamReader(dataStream))
                    {
                        // Read the content.Will be in this format
                        // vti_encoding:SR|utf8-nl
                        // vti_extenderversion: SR | 15.0.0.4505

                        string version = reader.ReadToEnd().Split('|')[2].Trim();

                        // Only compare the first three digits
                        var compareToVersion = new Version(minimallyRequiredVersion.Major, minimallyRequiredVersion.Minor, minimallyRequiredVersion.Build, 0);
                        hasMinimalVersion = new Version(version.Split('.')[0].ToInt32(), 0, version.Split('.')[3].ToInt32(), 0).CompareTo(compareToVersion) >= 0;
                    }
                }
            }
            catch (WebException ex)
            {
                Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ClientContextExtensions_HasMinimalServerLibraryVersion_Error, ex.ToDetailedString(clientContext));
            }
#endif
            return hasMinimalVersion;
        }

        /// <summary>
        /// Returns the name of the method calling ExecuteQueryRetry and ExecuteQueryRetryAsync
        /// </summary>
        /// <returns>A string with the method name</returns>
        private static string GetCallingPnPMethod()
        {
            StackTrace t = new StackTrace();

            string pnpMethod = "";
            try
            {
                for (int i = 0; i < t.FrameCount; i++)
                {
                    var frame = t.GetFrame(i);
                    var frameName = frame.GetMethod().Name;
                    if (frameName.Equals("ExecuteQueryRetry") || frameName.Equals("ExecuteQueryRetryAsync"))
                    {
                        var method = t.GetFrame(i + 1).GetMethod();

                        // Only return the calling method in case ExecuteQueryRetry was called from inside the PnP core library
                        if (method.Module.Name.Equals("OfficeDevPnP.Core.dll", StringComparison.InvariantCultureIgnoreCase))
                        {
                            pnpMethod = method.Name;
                        }
                        break;
                    }
                }
            }
            catch
            {
                // ignored
            }

            return pnpMethod;
        }

        /// <summary>
        /// Returns the request digest from the current session/site
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public static async Task<string> GetRequestDigest(this ClientContext context)
        {
            await new SynchronizationContextRemover();

            //InitializeSecurity(context);

            using (var handler = new HttpClientHandler())
            {
                string responseString = string.Empty;
                var accessToken = context.GetAccessToken();

                context.Web.EnsureProperty(w => w.Url);

                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = String.Format("{0}/_api/contextinfo", context.Web.Url);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }

                    HttpResponseMessage response = await httpClient.SendAsync(request);

                    if (response.IsSuccessStatusCode)
                    {
                        responseString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
                var contextInformation = JsonConvert.DeserializeObject<dynamic>(responseString);

                string formDigestValue = contextInformation.d.GetContextWebInformation.FormDigestValue;
                return await Task.Run(() => formDigestValue);
            }
        }

        private static void Context_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            if (!String.IsNullOrEmpty(e.WebRequestExecutor.RequestHeaders.Get("Authorization")))
            {
                accessToken = e.WebRequestExecutor.RequestHeaders.Get("Authorization").Replace("Bearer ", "");
            }
        }

#if !ONPREMISES
        /// <summary>
        /// BETA: Creates a Communication Site Collection
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="siteCollectionCreationInformation"></param>
        /// <returns></returns>
        public static async Task<ClientContext> CreateSiteAsync(this ClientContext clientContext, CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.CreateAsync(clientContext, siteCollectionCreationInformation);
        }

        /// <summary>
        /// BETA: Creates a Team Site Collection with no group
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="siteCollectionCreationInformation"></param>
        /// <returns></returns>
        public static async Task<ClientContext> CreateSiteAsync(this ClientContext clientContext, TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.CreateAsync(clientContext, siteCollectionCreationInformation);
        }

        /// <summary>
        /// BETA: Creates a Team Site Collection
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="siteCollectionCreationInformation"></param>
        /// <returns></returns>
        public static async Task<ClientContext> CreateSiteAsync(this ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.CreateAsync(clientContext, siteCollectionCreationInformation);
        }

        /// <summary>
        /// BETA: Groupifies a classic Team Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext instance of the site to be groupified</param>
        /// <param name="siteCollectionGroupifyInformation">Information needed to groupify this site</param>
        /// <returns>The clientcontext of the groupified site</returns>
        public static async Task<ClientContext> GroupifySiteAsync(this ClientContext clientContext, TeamSiteCollectionGroupifyInformation siteCollectionGroupifyInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.GroupifyAsync(clientContext, siteCollectionGroupifyInformation);
        }

        /// <summary>
        /// Checks if an alias is already used for an office 365 group or not
        /// </summary>
        /// <param name="clientContext">ClientContext of the site to operate against</param>
        /// <param name="alias">Alias to verify</param>
        /// <returns>True if in use, false otherwise</returns>
        public static async Task<bool> AliasExistsAsync(this ClientContext clientContext, string alias)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.AliasExistsAsync(clientContext, alias);
        }

        /// <summary>
        /// Enable MS Teams team on a group connected team site
        /// </summary>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        public static async Task<string> TeamifyAsync(this ClientContext clientContext)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.TeamifySiteAsync(clientContext);
        }
#endif
    }
}
