using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Security;
using System.Net;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
using System.Security.Cryptography.X509Certificates;

namespace OfficeDevPnP.Core.Tests
{
    static class TestCommon
    {
        #region Constructor
        static TestCommon()
        {
            // Read configuration data
            TenantUrl = ConfigurationManager.AppSettings["SPOTenantUrl"];
            DevSiteUrl = ConfigurationManager.AppSettings["SPODevSiteUrl"];            

#if !ONPREMISES
            if (string.IsNullOrEmpty(TenantUrl))
            {
                throw new ConfigurationErrorsException("Tenant site Url in App.config are not set up.");
            }
#endif
            if (string.IsNullOrEmpty(DevSiteUrl))
            {
                throw new ConfigurationErrorsException("Dev site url in App.config are not set up.");
            }



            // Trim trailing slashes
            TenantUrl = TenantUrl.TrimEnd(new[] { '/' });
            DevSiteUrl = DevSiteUrl.TrimEnd(new[] { '/' });

            if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOCredentialManagerLabel"]))
            {
                var tempCred = Core.Utilities.CredentialManager.GetCredential(ConfigurationManager.AppSettings["SPOCredentialManagerLabel"]);

                // username in format domain\user means we're testing in on-premises
                if (tempCred.UserName.IndexOf("\\") > 0)
                {
                    string[] userParts = tempCred.UserName.Split('\\');
                    Credentials = new NetworkCredential(userParts[1], tempCred.SecurePassword, userParts[0]);
                }
                else
                {
                    Credentials = new SharePointOnlineCredentials(tempCred.UserName, tempCred.SecurePassword);
                }
            }
            else
            {
                if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOUserName"]) &&
                    !String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOPassword"]))
                {
                    UserName = ConfigurationManager.AppSettings["SPOUserName"];
                    var password = ConfigurationManager.AppSettings["SPOPassword"];

                    Password = GetSecureString(password);
                    Credentials = new SharePointOnlineCredentials(UserName, Password);
                }
                else if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremUserName"]) &&
                         !String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremDomain"]) &&
                         !String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremPassword"]))
                {
                    Password = GetSecureString(ConfigurationManager.AppSettings["OnPremPassword"]);
                    Credentials = new NetworkCredential(ConfigurationManager.AppSettings["OnPremUserName"], Password, ConfigurationManager.AppSettings["OnPremDomain"]);
                }
                else if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["AppId"]) &&
                         !String.IsNullOrEmpty(ConfigurationManager.AppSettings["AppSecret"]))
                {
                    AppId = ConfigurationManager.AppSettings["AppId"];
                    AppSecret = ConfigurationManager.AppSettings["AppSecret"];
                }
                else if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["AppId"]) &&
                        !String.IsNullOrEmpty(ConfigurationManager.AppSettings["HighTrustIssuerId"]))
                {
                    AppId = ConfigurationManager.AppSettings["AppId"];
                    HighTrustCertificatePassword = ConfigurationManager.AppSettings["HighTrustCertificatePassword"];
                    HighTrustCertificatePath = ConfigurationManager.AppSettings["HighTrustCertificatePath"];
                    HighTrustIssuerId = ConfigurationManager.AppSettings["HighTrustIssuerId"];

                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["HighTrustCertificateStoreName"]))
                    {
                        StoreName result;
                        if (Enum.TryParse(ConfigurationManager.AppSettings["HighTrustCertificateStoreName"], out result))
                        {
                            HighTrustCertificateStoreName = result;
                        }
                    }
                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["HighTrustCertificateStoreLocation"]))
                    {
                        StoreLocation result;
                        if (Enum.TryParse(ConfigurationManager.AppSettings["HighTrustCertificateStoreLocation"], out result))
                        {
                            HighTrustCertificateStoreLocation = result;
                        }
                    }
                    HighTrustCertificateStoreThumbprint = ConfigurationManager.AppSettings["HighTrustCertificateStoreThumbprint"].Replace(" ", string.Empty);
                }
                else
                {
                    throw new ConfigurationErrorsException("Tenant credentials in App.config are not set up.");
                }
            }
        }
        #endregion

        #region Properties
        public static string TenantUrl { get; set; }
        public static string DevSiteUrl { get; set; }
        static string UserName { get; set; }
        static SecureString Password { get; set; }
        public static ICredentials Credentials { get; set; }
        public static string AppId { get; set; }
        static string AppSecret { get; set; }

        /// <summary>
        /// The path to the PFX file for the High Trust
        /// </summary>
        public static String HighTrustCertificatePath { get; set; }

        /// <summary>
        /// The password of the PFX file for the High Trust
        /// </summary>
        public static String HighTrustCertificatePassword { get; set; }

        /// <summary>
        /// The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer
        /// </summary>
        public static String HighTrustIssuerId { get; set; }

        /// <summary>
        /// The name of the store in the Windows certificate store where the High Trust certificate is stored
        /// </summary>
        public static StoreName? HighTrustCertificateStoreName { get; set; }

        /// <summary>
        /// The location of the High Trust certificate in the Windows certificate store
        /// </summary>
        public static StoreLocation? HighTrustCertificateStoreLocation { get; set; }

        /// <summary>
        /// The thumbprint / hash of the High Trust certificate in the Windows certificate store
        /// </summary>
        public static string HighTrustCertificateStoreThumbprint { get; set; }

        public static string TestWebhookUrl
        {
            get
            {
                return ConfigurationManager.AppSettings["WebHookTestUrl"];
            }
        }

        public static String AzureStorageKey
        {
            get
            {
                return ConfigurationManager.AppSettings["AzureStorageKey"];
            }
        }
        public static String TestAutomationDatabaseConnectionString
        {
            get
            {
                return ConfigurationManager.AppSettings["TestAutomationDatabaseConnectionString"];
            }
        }
        public static String AzureADCertPfxPassword
        {
            get
            {
                return ConfigurationManager.AppSettings["AzureADCertPfxPassword"];
            }
        }
        public static String AzureADClientId
        {
            get
            {
                return ConfigurationManager.AppSettings["AzureADClientId"];
            }
        }
        public static String NoScriptSite
        {
            get
            {
                return ConfigurationManager.AppSettings["NoScriptSite"];
            }
        }
        public static String ScriptSite
        {
            get
            {
                return ConfigurationManager.AppSettings["ScriptSite"];
            }
        }
        #endregion

        #region Methods
        public static ClientContext CreateClientContext()
        {
            return CreateContext(DevSiteUrl, Credentials);
        }

        public static ClientContext CreateClientContext(string url)
        {
            return CreateContext(url, Credentials);
        }

        public static ClientContext CreateTenantClientContext()
        {
            return CreateContext(TenantUrl, Credentials);
        }

        public static PnPClientContext CreatePnPClientContext(int retryCount = 10, int delay = 500)
        {
            PnPClientContext context;
            if (!String.IsNullOrEmpty(AppId) && !String.IsNullOrEmpty(AppSecret))
            {
                AuthenticationManager am = new AuthenticationManager();
                ClientContext clientContext = null;

                if (new Uri(DevSiteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    //clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, Core.Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(DevSiteUrl)), AppId, AppSecret, acsHostUrl: "windows-ppe.net", globalEndPointPrefix: "login");
                    clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret, AzureEnvironment.PPE);
                }
                else if (new Uri(DevSiteUrl).DnsSafeHost.Contains("sharepoint.de"))
                {
                    clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret, AzureEnvironment.Germany);
                }
                else
                {
                    clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret);
                }
                context = PnPClientContext.ConvertFrom(clientContext, retryCount, delay);
            }
            else
            {
                context = new PnPClientContext(DevSiteUrl, retryCount, delay);
                context.Credentials = Credentials;
            }

            context.RequestTimeout = Timeout.Infinite;
            return context;
        }


        public static bool AppOnlyTesting()
        {
            if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["AppId"]) &&
                !String.IsNullOrEmpty(ConfigurationManager.AppSettings["AppSecret"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOCredentialManagerLabel"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOUserName"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOPassword"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremUserName"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremDomain"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremPassword"]))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool TestAutomationSQLDatabaseAvailable()
        {
            string connectionString = TestAutomationDatabaseConnectionString;
            if (!String.IsNullOrEmpty(connectionString))
            {
                try
                {
                    var con = new SqlConnectionStringBuilder(connectionString);
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        return (conn.State == ConnectionState.Open);
                    }
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        private static ClientContext CreateContext(string contextUrl, ICredentials credentials)
        {
            ClientContext context;
            if (!String.IsNullOrEmpty(AppId) && !String.IsNullOrEmpty(AppSecret))
            {
                AuthenticationManager am = new AuthenticationManager();

                if (new Uri(DevSiteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    context = am.GetAppOnlyAuthenticatedContext(contextUrl, Core.Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(DevSiteUrl)), AppId, AppSecret, acsHostUrl: "windows-ppe.net", globalEndPointPrefix: "login");
                }
                else
                {
                    context = am.GetAppOnlyAuthenticatedContext(contextUrl, AppId, AppSecret);
                }
            }
            else
            {
                context = new ClientContext(contextUrl);
                context.Credentials = credentials;
            }

            context.RequestTimeout = Timeout.Infinite;
            return context;
        }

        private static SecureString GetSecureString(string input)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("Input string is empty and cannot be made into a SecureString", "input");

            var secureString = new SecureString();
            foreach (char c in input.ToCharArray())
                secureString.AppendChar(c);

            return secureString;
        }
        #endregion
    }
}
