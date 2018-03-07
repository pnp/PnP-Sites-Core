using Microsoft.SharePoint.Client;
using System;
using System.Reflection;

namespace OfficeDevPnP.Core
{
    /// <summary>
    /// Class that deals with PnPClientContext methods
    /// </summary>
    public class PnPClientContext : ClientContext
    {
        public int RetryCount { get; set; }
        public int Delay { get; set; }

        /// <summary>
        /// Converts ClientContext into PnPClientContext
        /// </summary>
        /// <param name="clientContext">A SharePoint ClientContext for resource operations</param>
        /// <param name="retryCount">Maximum amount of retries before giving up</param>
        /// <param name="delay">Initial delay in milliseconds</param>
        /// <returns></returns>
        public static PnPClientContext ConvertFrom(ClientContext clientContext, int retryCount = 10, int delay = 500)
        {
            var context = new PnPClientContext(clientContext.Url, retryCount, delay);

            context.AuthenticationMode = clientContext.AuthenticationMode;

            // In case of using networkcredentials in on premises or SharePointOnlineCredentials in Office 365
            if (clientContext.Credentials != null)
            {
                context.Credentials = clientContext.Credentials;
            }
            else
            {
                //Take over the form digest handling setting
                context.FormDigestHandlingEnabled = clientContext.FormDigestHandlingEnabled;

                // In case of app only or SAML
                context.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                    // the new delegate method
                    MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                    object[] parametersArray = new object[] { webRequestEventArgs };
                    methodInfo.Invoke(clientContext, parametersArray);
                };
            }

            return context;
        }

        /// <summary>
        /// Creates a ClientContext allowing you to override the default retry and delay values of ExecuteQueryRetry
        /// </summary>
        /// <param name="url">A SharePoint site URL</param>
        /// <param name="retryCount">Maximum amount of retries before giving up</param>
        /// <param name="delay">Initial delay in milliseconds</param>
        public PnPClientContext(string url, int retryCount = 10, int delay = 500) : base(url)
        {
            RetryCount = retryCount;
            Delay = delay;
        }

        /// <summary>
        /// Creates a ClientContext allowing you to override the default retry and delay values of ExecuteQueryRetry
        /// </summary>
        /// <param name="uri">A SharePoint site/web full URL</param>
        /// <param name="retryCount">Maximum amount of retries before giving up</param>
        /// <param name="delay">Initial delay in milliseconds</param>
        public PnPClientContext(Uri uri, int retryCount = 10, int delay = 500) : base(uri)
        {
            RetryCount = retryCount;
            Delay = delay;
        }

        /// <summary>
        /// Clones a PnPClientContext object while "taking over" the security context of the existing PnPClientContext instance
        /// </summary>
        /// <param name="siteUrl">Site URL to be used for cloned ClientContext</param>
        /// <returns>A PnPClientContext object created for the passed site URL</returns>
        public PnPClientContext Clone(string siteUrl)
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUrl));
            }
            return Clone(new Uri(siteUrl));
        }

        /// <summary>
        /// Clones a PnPClientContext object while "taking over" the security context of the existing PnPClientContext instance
        /// </summary>
        /// <param name="siteUri">Site URL to be used for cloned ClientContext</param>
        /// <returns>A PnPClientContext object created for the passed site URL</returns>
        public PnPClientContext Clone(Uri siteUri)
        {
            if (siteUri == null)
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUri));
            }

            var clonedClientContext = new PnPClientContext(siteUri)
            {
                RetryCount = this.RetryCount,
                Delay = this.Delay,
                AuthenticationMode = this.AuthenticationMode
            };


            // In case of using networkcredentials in on premises or SharePointOnlineCredentials in Office 365
            if (this.Credentials != null)
            {
                clonedClientContext.Credentials = this.Credentials;
            }
            else
            {
                //Take over the form digest handling setting
                clonedClientContext.FormDigestHandlingEnabled = this.FormDigestHandlingEnabled;

                // In case of app only or SAML
                clonedClientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                    // the new delegate method
                    var methodInfo = this.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                    var parametersArray = new object[] { webRequestEventArgs };
                    methodInfo.Invoke(this, parametersArray);
                };
            }

            return clonedClientContext;
        }

    }
}
