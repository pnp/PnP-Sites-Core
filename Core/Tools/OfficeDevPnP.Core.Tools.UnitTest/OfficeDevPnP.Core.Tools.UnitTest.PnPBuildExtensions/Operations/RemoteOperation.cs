using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Operations
{
    /// <summary>
    /// Used to differentiate the authentication scheme details during execution
    /// </summary>
    public enum AuthenticationType
    {
        DefaultCredentials,
        NetworkCredentials,
        Office365
    }

    public abstract class RemoteOperation
    {
        #region Construction
        public RemoteOperation(string targetUrl, AuthenticationType authType, string user, SecureString password, string AppInstanceId, string domain = "")
        {
            TargetSiteUrl = targetUrl;
            AuthType = authType;
            User = user;
            Password = password;
            Domain = domain;
            this.AppInstanceId = AppInstanceId;
        }
        #endregion

        #region Properties
        public string TargetSiteUrl { get; set; }
        public AuthenticationType AuthType { get; set; }
        public string User { get; set; }
        public SecureString Password { get; set; }
        public string Domain { get; set; }

        public string AppInstanceId { get; set; }

        public abstract string OperationPageUrl { get; }

        public Dictionary<string, string> PostParameters = new Dictionary<string, string>();
        #endregion

        #region Methods
        public bool Execute(string environment)
        {
            try
            {
                string page = GetRequest();
                AnalyzeRequestResponse(page);
                if (environment.ToLower().Equals("online"))
                {
                    SetPostVariablesOnline();
                    Thread.Sleep(1000);
                }
                else
                {
                    SetPostVariablesOnPremise();
                    Thread.Sleep(1000);
                }
                string PostRequestResult = MakePostRequest(page).ToString();
                if (PostRequestResult.Contains("The app identifier has been successfully created."))
                {
                    return true;
                }
                else if (PostRequestResult.Contains("Site Settings"))
                {
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                // TODO - Give better description for the exception
                throw new Exception("Execute failed.", ex);
            }
        }

        /// <summary>
        /// For optinal processing of the page in the inherited class
        /// </summary>
        /// <param name="page"></param>
        public virtual void AnalyzeRequestResponse(string page)
        {

        }

        /// <summary>
        /// To be implemented based on usage scenario
        /// </summary>
        public abstract void SetPostVariablesOnline();
        public abstract void SetPostVariablesOnPremise();

        /// <summary>
        /// Handels the initial request of the page using given identity
        /// </summary>
        /// <returns></returns>
        protected string GetRequest()
        {
            string returnString = string.Empty;
            string url = RemoteOperationUtilities.FormatOperationUrlString(TargetSiteUrl, OperationPageUrl);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

            //Set auth options based on auth model
            ModifyRequestBasedOnAuthPattern(request);

            // Set some reasonable limits on resources used by this request
            request.MaximumAutomaticRedirections = 6;
            request.MaximumResponseHeadersLength = 6;
            // Set user agent as valid text
            request.UserAgent = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)";

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                Encoding encode = System.Text.Encoding.GetEncoding("utf-8");

                Stream responseStream = response.GetResponseStream();
                if (response.ContentEncoding.ToLower().Contains("gzip"))
                {
                    responseStream = new GZipStream(responseStream, CompressionMode.Decompress);
                }
                else if (response.ContentEncoding.ToLower().Contains("deflate"))
                {
                    responseStream = new DeflateStream(responseStream, CompressionMode.Decompress);
                }

                // Get the response stream and store that as string
                using (StreamReader reader = new StreamReader(responseStream, encode))
                {
                    returnString = reader.ReadToEnd();
                }

                return returnString;
            }

        }

        /// <summary>
        /// Used to modify the HTTP request based on authentication type
        /// </summary>
        /// <param name="request"></param>
        private void ModifyRequestBasedOnAuthPattern(HttpWebRequest request)
        {

            // Change the model based on used auth type
            switch (AuthType)
            {
                case AuthenticationType.DefaultCredentials:
                    request.Credentials = CredentialCache.DefaultCredentials;
                    break;
                case AuthenticationType.NetworkCredentials:
                    NetworkCredential credential = new NetworkCredential(User, Password, Domain);
                    CredentialCache credentialCache = new CredentialCache();
                    credentialCache.Add(new Uri(TargetSiteUrl), "NTLM", credential);
                    request.Credentials = credentialCache;
                    break;
                case AuthenticationType.Office365:
                    SharePointOnlineCredentials Credentials = new SharePointOnlineCredentials(User, this.Password);
                    Uri tenantUrlUri = new Uri(TargetSiteUrl);
                    string authCookieValue = Credentials.GetAuthenticationCookie(tenantUrlUri);
                    // Create fed auth Cookie and set that to http request properly to access Office365 site
                    Cookie fedAuth = new Cookie()
                    {
                        Name = "SPOIDCRL",
                        Value = authCookieValue.TrimStart("SPOIDCRL=".ToCharArray()),
                        Path = "/",
                        Secure = true,
                        HttpOnly = true,
                        Domain = new Uri(TargetSiteUrl).Host
                    };
                    // Hookup authentication cookie to request
                    request.CookieContainer = new CookieContainer();
                    request.CookieContainer.Add(fedAuth);
                    break;
                default:

                    break;
            }
        }

        /// <summary>
        /// Responsible of accessing the page and submitting the post. Generic handler for the post access
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        protected string MakePostRequest(string webPage)
        {
            // Add required stuff to validate the page request
            string postBody = "__REQUESTDIGEST=" + RemoteOperationUtilities.ReadHiddenField(webPage, "__REQUESTDIGEST") +
                                "&__EVENTVALIDATION=" + RemoteOperationUtilities.ReadHiddenField(webPage, "__EVENTVALIDATION") +
                                "&__VIEWSTATE=" + RemoteOperationUtilities.ReadHiddenField(webPage, "__VIEWSTATE");

            // Add operation specific parameters
            foreach (var item in PostParameters)
            {
                postBody = postBody + string.Format("&{0}={1}", item.Key, item.Value);
            }

            string results = string.Empty;

            try
            {
                string url = RemoteOperationUtilities.FormatOperationUrlString(TargetSiteUrl, OperationPageUrl);

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                if (AuthType != AuthenticationType.Office365)
                {
                    // Get X-RequestDigest header for post. Required for most of the operations
                    request.Headers.Add("X-RequestDigest", GetUpdatedFormDigest(TargetSiteUrl));
                }

                // Note that this assumes that we use particular identity running the thread
                ModifyRequestBasedOnAuthPattern(request);
                Thread.Sleep(3000);

                // Set some reasonable limits on resources used by this request
                request.MaximumAutomaticRedirections = 6;
                request.MaximumResponseHeadersLength = 6;
                // Set credentials to use for this request.
                request.ContentType = "application/x-www-form-urlencoded";

                byte[] postByte = Encoding.UTF8.GetBytes(postBody);
                request.ContentLength = postByte.Length;
                Stream postStream = request.GetRequestStream();
                Thread.Sleep(3000);
                postStream.Write(postByte, 0, postByte.Length);
                postStream.Close();

                HttpWebResponse wResp = (HttpWebResponse)request.GetResponse();
                postStream = wResp.GetResponseStream();
                StreamReader postReader = new StreamReader(postStream);

                results = postReader.ReadToEnd();

                postReader.Close();
            }
            catch (Exception ex)
            {
                // Give better description for the exception
                throw new Exception("MakePostRequest failed.", ex);
            }

            return results;
        }

        /// <summary>
        /// Used to get updated form digest for the operation. Required for most of the operations in on-prem or with Office365-D
        /// </summary>
        /// <param name="siteUrl">Url to access</param>
        /// <returns></returns>
        private string GetUpdatedFormDigest(string siteUrl)
        {
            try
            {
                string url = siteUrl.TrimEnd(new char[] { '/' }) + "/_vti_bin/sites.asmx";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";

                // Note that this assumes that we use particular identity running the thread
                request.Credentials = CredentialCache.DefaultCredentials;

                // Set some reasonable limits on resources used by this request
                request.MaximumAutomaticRedirections = 6;
                request.MaximumResponseHeadersLength = 6;
                // Set credentials to use for this request.
                request.ContentType = "text/xml";

                var payload =
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "  <soap:Body>" +
                    "    <GetUpdatedFormDigest xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\" />" +
                    "  </soap:Body>" +
                    "</soap:Envelope>";

                byte[] postByte = Encoding.UTF8.GetBytes(payload);
                request.ContentLength = postByte.Length;
                Stream postStream = request.GetRequestStream();
                postStream.Write(postByte, 0, postByte.Length);
                postStream.Close();

                HttpWebResponse wResp = (HttpWebResponse)request.GetResponse();
                postStream = wResp.GetResponseStream();
                StreamReader postReader = new StreamReader(postStream);

                string results = postReader.ReadToEnd();

                postReader.Close();

                var startTag = "<GetUpdatedFormDigestResult>";
                var endTag = "</GetUpdatedFormDigestResult>";
                var startTagIndex = results.IndexOf(startTag);
                var endTagIndex = results.IndexOf(endTag, startTagIndex + startTag.Length);
                string newFormDigest = null;
                if ((startTagIndex >= 0) && (endTagIndex > startTagIndex))
                {
                    newFormDigest = results.Substring(startTagIndex + startTag.Length, endTagIndex - startTagIndex - startTag.Length);
                }
                return newFormDigest;

            }
            catch (Exception ex)
            {
                // TODO - Give better description for the exception
                throw new Exception("GetUpdatedFormDigest failed.", ex);
            }
        }
        #endregion
    }
}
