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
        /// <param name="siteUrl">Site url to be used for cloned ClientContext</param>
        /// <returns>A ClientContext object created for the passed site url</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, string siteUrl)
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUrl));
            }

            return clientContext.Clone(new Uri(siteUrl));
        }


        /// <summary>
        /// Executes the current set of data retrieval queries and method invocations and retries it if needed.
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"></param>
        public static void ExecuteQueryRetry(this ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500, string userAgent = null)
        {
            ExecuteQueryImplementation(clientContext, retryCount, delay, userAgent);
        }

        private static void ExecuteQueryImplementation(ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500, string userAgent = null)
        {
            var clientTag = string.Empty;
            if (clientContext is PnPClientContext)
            {
                retryCount = (clientContext as PnPClientContext).RetryCount;
                delay = (clientContext as PnPClientContext).Delay;
                clientTag = (clientContext as PnPClientContext).ClientTag;
            }

            int retryAttempts = 0;
            int backoffInterval = delay;
            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");

            if (delay <= 0)
                throw new ArgumentException("Provide a delay greater than zero.");

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
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
                    clientContext.ClientTag = clientTag;

                    // Make CSOM request more reliable by disabling the return value cache. Given we 
                    // often clone context objects and the default value is
#if !ONPREMISES
                    clientContext.DisableReturnValueCache = true;
#elif SP2016
                    clientContext.DisableReturnValueCache = true;
#endif
                    // Add event handler to "insert" app decoration header to mark the PnP Sites Core library as a known application
                    EventHandler<WebRequestEventArgs> appDecorationHandler = (s, e) =>
                    {
                        if (string.IsNullOrEmpty(userAgent) && !string.IsNullOrEmpty(ClientContextExtensions.userAgentFromConfig))
                        {
                            userAgent = userAgentFromConfig;
                        }

                        e.WebRequestExecutor.WebRequest.UserAgent = string.IsNullOrEmpty(userAgent) ? $"{PnPCoreUtilities.PnPCoreUserAgent}" : userAgent;
                    };

                    clientContext.ExecutingWebRequest += appDecorationHandler;

                    // DO NOT CHANGE THIS TO EXECUTEQUERYRETRY
                    clientContext.ExecuteQuery();

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

                        //Add delay for retry
                        Thread.Sleep(backoffInterval);

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
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="siteUrl">Site url to be used for cloned ClientContext</param>
        /// <returns>A ClientContext object created for the passed site url</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, Uri siteUrl)
        {
            if (siteUrl == null)
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUrl));
            }

            ClientContext clonedClientContext = new ClientContext(siteUrl);
            clonedClientContext.AuthenticationMode = clientContext.AuthenticationMode;
            clonedClientContext.ClientTag = clientContext.ClientTag;
#if !ONPREMISES
            clonedClientContext.DisableReturnValueCache = clientContext.DisableReturnValueCache;
#elif SP2016
            clonedClientContext.DisableReturnValueCache = clientContext.DisableReturnValueCache;
#endif


            // In case of using networkcredentials in on premises or SharePointOnlineCredentials in Office 365
            if (clientContext.Credentials != null)
            {
                clonedClientContext.Credentials = clientContext.Credentials;
            }
            else
            {
                //Take over the form digest handling setting
                clonedClientContext.FormDigestHandlingEnabled = (clientContext as ClientContext).FormDigestHandlingEnabled;

                // In case of app only or SAML
                clonedClientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                    // the new delegate method
                    MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                    object[] parametersArray = new object[] { webRequestEventArgs };
                    methodInfo.Invoke(clientContext, parametersArray);
                };
            }

            return clonedClientContext;
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
            if (clientContext.Credentials == null)
            {
                return true;
            }

            return false;
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

        private static string GetCallingPnPMethod()
        {
            StackTrace t = new StackTrace();

            string pnpMethod = "";
            try
            {
                for (int i = 0; i < t.FrameCount; i++)
                {
                    var frame = t.GetFrame(i);
                    if (frame.GetMethod().Name.Equals("ExecuteQueryRetry"))
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
            //InitializeSecurity(context);

            using (var handler = new HttpClientHandler())
            {
                string responseString = string.Empty;
                var accessToken = context.GetAccessToken();

                if (String.IsNullOrEmpty(accessToken))
                {
                    context.Web.EnsureProperty(w => w.Url);
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string requestUrl = String.Format("{0}/_api/contextinfo", context.Web.Url);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");
                    if(!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
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
            return await SiteCollection.CreateAsync(clientContext, siteCollectionCreationInformation);
        }
#endif
    }
}
