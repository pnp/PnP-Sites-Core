using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Resources;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions
{
    public class AppManager
    {
        #region Private variables
        private string sharePointUrl;
        private AuthenticationType authenticationType = AuthenticationType.Office365;
        private string userName = "";
        private SecureString password;
        #endregion

        #region Construction
        public AppManager(string sharePointUrl, string credentialManagerLabel)
        {
            this.sharePointUrl = sharePointUrl;
            
            if (sharePointUrl.IndexOf("sharepoint.com") > -1)
            {
                authenticationType = AuthenticationType.Office365;
            }
            else
            {
                authenticationType = AuthenticationType.NetworkCredentials;
            }

            NetworkCredential cred = CredentialManager.GetCredential(credentialManagerLabel);
            userName = cred.UserName;
            password = cred.SecurePassword;
        }

        public AppManager(string sharePointUrl, AuthenticationType authenticationType, string credentialManagerLabel)
        {
            this.sharePointUrl = sharePointUrl;
            this.authenticationType = authenticationType;
            NetworkCredential cred = CredentialManager.GetCredential(credentialManagerLabel);
            userName = cred.UserName;
            password = cred.SecurePassword;        
        }
        public AppManager(string sharePointUrl, string userName, SecureString password)
        {
            this.sharePointUrl = sharePointUrl;

            if (sharePointUrl.IndexOf("sharepoint.com") > -1)
            {
                authenticationType = AuthenticationType.Office365;
            }
            else
            {
                authenticationType = AuthenticationType.NetworkCredentials;
            }

            this.userName = userName;
            this.password = password;
        }

        public AppManager(string sharePointUrl, AuthenticationType authenticationType, string userName, SecureString password)
        {
            this.sharePointUrl = sharePointUrl;
            this.authenticationType = authenticationType;
            this.userName = userName;
            this.password = password;
        }
        #endregion

        #region Properties

        #endregion

        #region Public Methods
        public bool CreateAppPackageForProviderHostedApp(string sharePointProjectFile, string sharePointWebProjectFile, string clientId, string clientSecret, string applicationHost)
        {
            bool createAppPackageResult = false;

            if (String.IsNullOrEmpty(sharePointProjectFile) || !System.IO.File.Exists(sharePointProjectFile))
            {
                throw new ArgumentException(String.Format("Provide SharePoint project file ({0}) is invalid.", sharePointProjectFile));
            }

            if (String.IsNullOrEmpty(sharePointWebProjectFile) || !System.IO.File.Exists(sharePointWebProjectFile))
            {
                throw new ArgumentException(String.Format("Provide SharePoint Web project file ({0}) is invalid.", sharePointWebProjectFile));
            }

            if (String.IsNullOrEmpty(clientId))
            {
                throw new ArgumentException("Please provide a client id");
            }

            if (String.IsNullOrEmpty(clientSecret))
            {
                throw new ArgumentException("Please provide a client secret");
            }

            if (String.IsNullOrEmpty(applicationHost))
            {
                throw new ArgumentException("Please provide an application host");
            }

            // Do we already have a publishing XML defined?


            // Get a base template that will be used for the publishing
            using (Stream publishingTemplate = ResourceManager.GetPublishingXmlTemplate(true, true, PublishingTypes.AzureWebSite))
            {

            }




            return createAppPackageResult;
        }



        public bool RegisterApplication(string clientId, string clientSecret, string title, string appDomain, string redirectUri)
        {
            bool appRegNewResult = false;

            if (String.IsNullOrEmpty(sharePointUrl) || !ValidateUri(sharePointUrl))
            {
                throw new ArgumentException("Please provide a valid value for the SharePoint url");
            }

            if (String.IsNullOrEmpty(title))
            {
                throw new ArgumentException("Please provide a value for title");
            }

            if (String.IsNullOrEmpty(appDomain))
            {
                throw new ArgumentException("Please provide a value for appDomain");
            }

            if (String.IsNullOrEmpty(redirectUri) || !ValidateUri(redirectUri))
            {
                throw new ArgumentException("Please provide a valid value for the redirect uri");
            }

            if (clientId == null)
            {
                clientId = AppRegNew.GenerateAppId();
            }


            if (String.IsNullOrEmpty(clientSecret))
            {
                clientSecret = AppRegNew.GenerateAppSecret();
            }

            AppRegNew appRegNew = new AppRegNew(sharePointUrl, this.authenticationType, userName, password);
            appRegNew.AppId = clientId;
            appRegNew.AppSecret = clientSecret;
            appRegNew.Title = title;
            appRegNew.HostUri = appDomain;
            appRegNew.RedirectUri = redirectUri;
            appRegNewResult = appRegNew.Execute(this.authenticationType == AuthenticationType.Office365 ? "online" : "onpremises");

            Thread.Sleep(1000);
            if (appRegNewResult == true)
            {
                Console.WriteLine("App Registration done for \n \t App Id: " + appRegNew.AppId
                     + "\n \t App Secret: " + appRegNew.AppSecret
                     + "\n \t App Title: " + appRegNew.Title
                     + "\n \t Host Uri: " + appRegNew.HostUri
                     + "\n \t Redirect Uri: " + appRegNew.RedirectUri);
                Console.WriteLine("App Registered Successfully");
            }

            else
            {
                Console.WriteLine("App Registration failed for below app details \n \t App Id: " + appRegNew.AppId
                     + "\n \t App Secret: " + appRegNew.AppSecret
                     + "\n \t App Title: " + appRegNew.Title
                     + "\n \t Host Uri: " + appRegNew.HostUri
                     + "\n \t Redirect Uri: " + appRegNew.RedirectUri);
            }
            return appRegNewResult;
        }
        #endregion

        #region Private methods
        private bool ValidateUri(string url)
        {            
            url = url.Trim();

            if (url.Length == 0)
            {
                return false;
            }

            Uri uri;
            if (!Uri.TryCreate(url, UriKind.Absolute, out uri))
            {
                return false;
            }

            if (uri.Scheme != Uri.UriSchemeHttp &&
                uri.Scheme != Uri.UriSchemeHttps)
            {
                return false;
            }

            return true;
        }
        #endregion

    }
}
