using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
#if !NETSTANDARD2_0
using System.Windows.Forms;
#endif
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.IdentityModel.TokenProviders.ADFS;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Utilities.Async;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Utilities.Context;
using System.Web;

namespace OfficeDevPnP.Core
{
    /// <summary>
    /// Enum to identify the supported Office 365 hosting environments
    /// </summary>
    public enum AzureEnvironment
    {
        Production=0,
        PPE=1,
        China=2,
        Germany=3,
        USGovernment=4
    }

    /// <summary>
    /// This manager class can be used to obtain a SharePointContext object
    /// </summary>
    ///
    public class AuthenticationManager
    {
        private const string SHAREPOINT_PRINCIPAL = "00000003-0000-0ff1-ce00-000000000000";

        private SharePointOnlineCredentials sharepointOnlineCredentials;
        private string appOnlyAccessToken;
        private string azureADCredentialsToken;
        private object tokenLock = new object();
        private CookieContainer fedAuth = null;
        private string _contextUrl;
        private TokenCache _tokenCache;
        private string _commonAuthority = "https://login.windows.net/Common";
        private static AuthenticationContext _authContext = null;
        private string _clientId;
        private Uri _redirectUri;

        #region Construction
        public AuthenticationManager()
        {
#if !ONPREMISES
            // Set the TLS preference. Needed on some server os's to work when Office 365 removes support for TLS 1.0
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
#endif
        }
        #endregion


        #region Authenticating against SharePoint Online using credentials or app-only
        /// <summary>
        /// Returns a SharePointOnline ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="tenantUser">User to be used to instantiate the ClientContext object</param>
        /// <param name="tenantUserPassword">Password of the user used to instantiate the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetSharePointOnlineAuthenticatedContextTenant(string siteUrl, string tenantUser, string tenantUserPassword)
        {
            var spoPassword = Utilities.EncryptionUtility.ToSecureString(tenantUserPassword);
            return GetSharePointOnlineAuthenticatedContextTenant(siteUrl, tenantUser, spoPassword);
        }

        /// <summary>
        /// Returns a SharePointOnline ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="tenantUser">User to be used to instantiate the ClientContext object</param>
        /// <param name="tenantUserPassword">Password (SecureString) of the user used to instantiate the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetSharePointOnlineAuthenticatedContextTenant(string siteUrl, string tenantUser, SecureString tenantUserPassword)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_GetContext, siteUrl);
            Log.Debug(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_TenantUser, tenantUser);

            if (sharepointOnlineCredentials == null)
            {
                sharepointOnlineCredentials = new SharePointOnlineCredentials(tenantUser, tenantUserPassword);
            }

            var ctx = new ClientContext(siteUrl);
            ctx.Credentials = sharepointOnlineCredentials;
#if !ONPREMISES || SP2016 || SP2019
            ctx.DisableReturnValueCache = true;
#endif

            return ctx;
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetAppOnlyAuthenticatedContext(string siteUrl, string appId, string appSecret)
        {
            return GetAppOnlyAuthenticatedContext(siteUrl, TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl)), appId, appSecret);
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetAppOnlyAuthenticatedContext(string siteUrl, string appId, string appSecret, AzureEnvironment environment = AzureEnvironment.Production)
        {
            return GetAppOnlyAuthenticatedContext(siteUrl, TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl)), appId, appSecret, GetAzureADACSEndPoint(environment), GetAzureADACSEndPointPrefix(environment));
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="realm">Realm of the environment (tenant) that requests the ClientContext object</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="acsHostUrl">Azure ACS host, defaults to accesscontrol.windows.net but internal pre-production environments use other hosts</param>
        /// <param name="globalEndPointPrefix">Azure ACS endpoint prefix, defaults to accounts but internal pre-production environments use other prefixes</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetAppOnlyAuthenticatedContext(string siteUrl, string realm, string appId, string appSecret, string acsHostUrl = "accesscontrol.windows.net", string globalEndPointPrefix = "accounts")
        {
            EnsureToken(siteUrl, realm, appId, appSecret, acsHostUrl, globalEndPointPrefix);
            ClientContext clientContext = Utilities.TokenHelper.GetClientContextWithAccessToken(siteUrl, appOnlyAccessToken);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.SharePointACSAppOnly,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                Realm = realm,
                ClientId = appId,
                ClientSecret = appSecret,
                AcsHostUrl = acsHostUrl,
                GlobalEndPointPrefix = globalEndPointPrefix
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

        /// <summary>
        /// Get's the Azure ASC login end point for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure ASC login endpoint</returns>
        public string GetAzureADACSEndPoint(AzureEnvironment environment)
        {
            switch (environment)
            {
                case AzureEnvironment.Production:
                    {
                        return "accesscontrol.windows.net";
                    }
                case AzureEnvironment.Germany:
                    {
                        return "microsoftonline.de";
                    }
                case AzureEnvironment.China:
                    {
                        return "accesscontrol.chinacloudapi.cn";
                    }
                case AzureEnvironment.USGovernment:
                    {
                        return "microsoftonline.us";
                    }
                case AzureEnvironment.PPE:
                    {
                        return "windows-ppe.net";
                    }
                default:
                    {
                        return "accesscontrol.windows.net";
                    }
            }
        }

        /// <summary>
        /// Get's the Azure ACS login end point prefix for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure ACS login endpoint prefix</returns>
        public string GetAzureADACSEndPointPrefix(AzureEnvironment environment)
        {
            switch (environment)
            {
                case AzureEnvironment.Production:
                    {
                        return "accounts";
                    }
                case AzureEnvironment.Germany:
                    {
                        return "login";
                    }
                case AzureEnvironment.China:
                    {
                        return "accounts";
                    }
                case AzureEnvironment.USGovernment:
                    {
                        return "login";
                    }
                case AzureEnvironment.PPE:
                    {
                        return "login";
                    }
                default:
                    {
                        return "accounts";
                    }
            }
        }

        /// <summary>
        /// Ensure that AppAccessToken is filled with a valid string representation of the OAuth AccessToken. This method will launch handle with token cleanup after the token expires
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="realm">Realm of the environment (tenant) that requests the ClientContext object</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="acsHostUrl">Azure ACS host, defaults to accesscontrol.windows.net but internal pre-production environments use other hosts</param>
        /// <param name="globalEndPointPrefix">Azure ACS endpoint prefix, defaults to accounts but internal pre-production environments use other prefixes</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private void EnsureToken(string siteUrl, string realm, string appId, string appSecret, string acsHostUrl, string globalEndPointPrefix)
        {
            if (appOnlyAccessToken == null)
            {
                lock (tokenLock)
                {
                    Log.Debug(Constants.LOGGING_SOURCE, "AuthenticationManager:EnsureToken(siteUrl:{0},realm:{1},appId:{2},appSecret:PRIVATE)", siteUrl, realm, appId);
                    if (appOnlyAccessToken == null)
                    {
                        Utilities.TokenHelper.Realm = realm;
                        Utilities.TokenHelper.ServiceNamespace = realm;
                        Utilities.TokenHelper.ClientId = appId;
                        Utilities.TokenHelper.ClientSecret = appSecret;

                        if (!String.IsNullOrEmpty(acsHostUrl))
                        {
                            Utilities.TokenHelper.AcsHostUrl = acsHostUrl;
                        }

                        if (globalEndPointPrefix != null)
                        {
                            Utilities.TokenHelper.GlobalEndPointPrefix = globalEndPointPrefix;
                        }

                        var response = Utilities.TokenHelper.GetAppOnlyAccessToken(SHAREPOINT_PRINCIPAL, new Uri(siteUrl).Authority, realm);
                        string token = response.AccessToken;
                        ThreadPool.QueueUserWorkItem(obj =>
                        {
                            try
                            {
                                Log.Debug(Constants.LOGGING_SOURCE, "Lease expiration date: {0}", response.ExpiresOn);
                                var lease = GetAccessTokenLease(response.ExpiresOn);
                                lease =
                                    TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
                                Thread.Sleep(lease);
                                appOnlyAccessToken = null;
                            }
                            catch (Exception ex)
                            {
                                Log.Warning(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManger_ProblemDeterminingTokenLease, ex);
                                appOnlyAccessToken = null;
                            }
                        });
                        appOnlyAccessToken = token;
                    }
                }
            }
        }

        /// <summary>
        /// Get the access token lease time span.
        /// </summary>
        /// <param name="expiresOn">The ExpiresOn time of the current access token</param>
        /// <returns>Returns a TimeSpan represents the time interval within which the current access token is valid thru.</returns>
        private TimeSpan GetAccessTokenLease(DateTime expiresOn)
        {
            DateTime now = DateTime.UtcNow;
            DateTime expires = expiresOn.Kind == DateTimeKind.Utc ?
                expiresOn : TimeZoneInfo.ConvertTimeToUtc(expiresOn);
            TimeSpan lease = expires - now;
            return lease;
        }

#if !NETSTANDARD2_0
        /// <summary>
        /// Returns a SharePoint on-premises / SharePoint Online ClientContext object. Requires claims based authentication with FedAuth cookie.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="icon">Optional icon to use for the popup form</param>
        /// <param name="scriptErrorsSuppressed">Optional parameter to set WebBrowser.ScriptErrorsSuppressed value in the popup form</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetWebLoginClientContext(string siteUrl, System.Drawing.Icon icon = null, bool scriptErrorsSuppressed = true)
        {
            var authCookiesContainer = new CookieContainer();
            var siteUri = new Uri(siteUrl);

            var thread = new Thread(() =>
            {
                var form = new System.Windows.Forms.Form();
                if (icon != null)
                {
                    form.Icon = icon;
                }
                var browser = new System.Windows.Forms.WebBrowser
                {
                    ScriptErrorsSuppressed = scriptErrorsSuppressed,
                    Dock = DockStyle.Fill
                };

                form.SuspendLayout();
                form.Width = 900;
                form.Height = 500;
                form.Text = $"Log in to {siteUrl}";
                form.Controls.Add(browser);
                form.ResumeLayout(false);

                browser.Navigate(siteUri);

                browser.Navigated += (sender, args) =>
                {
                    if (siteUri.Host.Equals(args.Url.Host))
                    {
                        var cookieString = CookieReader.GetCookie(siteUrl).Replace("; ", ",").Replace(";", ",");

                        // Get FedAuth and rtFa cookies issued by ADFS when accessing claims aware applications.
                        // - or get the EdgeAccessCookie issued by the Web Application Proxy (WAP) when accessing non-claims aware applications (Kerberos).
                        IEnumerable<string> authCookies = null;
                        if (Regex.IsMatch(cookieString, "FedAuth", RegexOptions.IgnoreCase))
                        {
                            authCookies = cookieString.Split(',').Where(c => c.StartsWith("FedAuth", StringComparison.InvariantCultureIgnoreCase) || c.StartsWith("rtFa", StringComparison.InvariantCultureIgnoreCase));
                        } else if (Regex.IsMatch(cookieString, "EdgeAccessCookie", RegexOptions.IgnoreCase))
                        {
                            authCookies = cookieString.Split(',').Where(c => c.StartsWith("EdgeAccessCookie", StringComparison.InvariantCultureIgnoreCase));
                        }
                        if (authCookies != null)
                        {
                            authCookiesContainer.SetCookies(siteUri, string.Join(",", authCookies));
                            form.Close();
                        }
                    }
                };

                form.Focus();
                form.ShowDialog();
                browser.Dispose();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (authCookiesContainer.Count > 0)
            {
                var ctx = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
                ctx.DisableReturnValueCache = true;
#endif
                ctx.ExecutingWebRequest += (sender, e) => e.WebRequestExecutor.WebRequest.CookieContainer = authCookiesContainer;
                return ctx;
            }

            return null;
        }
#endif
#endregion

#region Authenticating against SharePoint on-premises using credentials
        /// <summary>
        /// Returns a SharePoint on-premises / SharePoint Online Dedicated ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="user">User to be used to instantiate the ClientContext object</param>
        /// <param name="password">Password of the user used to instantiate the ClientContext object</param>
        /// <param name="domain">Domain of the user used to instantiate the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetNetworkCredentialAuthenticatedContext(string siteUrl, string user, string password, string domain)
        {
            ClientContext clientContext = new ClientContext(siteUrl)
            {
#if !ONPREMISES || SP2016 || SP2019
                DisableReturnValueCache = true,
#endif
                Credentials = new NetworkCredential(user, password, domain)
            };
            return clientContext;
        }

        /// <summary>
        /// Returns a SharePoint on-premises / SharePoint Online Dedicated ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="user">User to be used to instantiate the ClientContext object</param>
        /// <param name="password">Password (SecureString) of the user used to instantiate the ClientContext object</param>
        /// <param name="domain">Domain of the user used to instantiate the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetNetworkCredentialAuthenticatedContext(string siteUrl, string user, SecureString password, string domain)
        {
            ClientContext clientContext = new ClientContext(siteUrl)
            {
#if !ONPREMISES || SP2016 || SP2019
                DisableReturnValueCache = true,
#endif
                Credentials = new NetworkCredential(user, password, domain)
            };
            return clientContext;
        }

#if !NETSTANDARD2_0
        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="certificatePath">Full path to the private key certificate (.pfx) used to authenticate</param>
        /// <param name="certificatePassword">Password used for the private key certificate (.pfx)</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppOnlyAuthenticatedContext(string siteUrl, string clientId, string certificatePath, string certificatePassword, string certificateIssuerId)
        {
            var certPassword = Utilities.EncryptionUtility.ToSecureString(certificatePassword);
            return GetHighTrustCertificateAppOnlyAuthenticatedContext(siteUrl, clientId, certificatePath, certPassword, certificateIssuerId);
        }
#endif

#if !NETSTANDARD2_0
        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="certificatePath">Full path to the private key certificate (.pfx) used to authenticate</param>
        /// <param name="certificatePassword">Password used for the private key certificate (.pfx)</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppOnlyAuthenticatedContext(string siteUrl, string clientId, string certificatePath, SecureString certificatePassword, string certificateIssuerId)
        {
            var certificate = new X509Certificate2(certificatePath, certificatePassword);
            return GetHighTrustCertificateAppOnlyAuthenticatedContext(siteUrl, clientId, certificate, certificateIssuerId);
        }
#endif

#if !NETSTANDARD2_0
        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="storeName">The name of the store for the certificate</param>
        /// <param name="storeLocation">The location of the store for the certificate</param>
        /// <param name="thumbPrint">The thumbprint of the certificate to locate in the store</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppOnlyAuthenticatedContext(string siteUrl, string clientId, StoreName storeName, StoreLocation storeLocation, string thumbPrint, string certificateIssuerId)
        {
            // Retrieve the certificate from the Windows Certificate Store
            var cert = Utilities.X509CertificateUtility.LoadCertificate(storeName, storeLocation, thumbPrint);
            return GetHighTrustCertificateAppOnlyAuthenticatedContext(siteUrl, clientId, cert, certificateIssuerId);
        }
#endif

#if !NETSTANDARD2_0
        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="certificate">Private key certificate (.pfx) used to authenticate</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppOnlyAuthenticatedContext(string siteUrl, string clientId, X509Certificate2 certificate, string certificateIssuerId)
        {
            var siteUri = new Uri(siteUrl);
            var clientContext = new ClientContext(siteUri);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif

            // Feed the TokenHelper the SharePoint information so it doesn't try to fetch it from the config file
            TokenHelper.Realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            TokenHelper.ClientId = clientId;
            TokenHelper.ClientCertificate = certificate;
            TokenHelper.IssuerId = certificateIssuerId;

            // Configure the handler to generate the Bearer Access Token on each request and add it to the request
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                var accessToken = TokenHelper.GetS2SAccessTokenWithWindowsIdentity(siteUri, null);
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return clientContext;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="certificatePath">Full path to the private key certificate (.pfx) used to authenticate</param>
        /// <param name="certificatePassword">Password used for the private key certificate (.pfx)</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <param name="loginName">
        /// Name of the user (login name) on whose behalf to create the access token.
        /// Supported input formats are SID and User Principal Name (UPN).
        /// If the parameter is left empty (including null) an App Only Context will be created.
        /// </param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppAuthenticatedContext(string siteUrl, string clientId, string certificatePath, string certificatePassword, string certificateIssuerId, string loginName)
        {
            var certPassword = Utilities.EncryptionUtility.ToSecureString(certificatePassword);
            return GetHighTrustCertificateAppAuthenticatedContext(siteUrl, clientId, certificatePath, certPassword, certificateIssuerId, loginName);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="certificatePath">Full path to the private key certificate (.pfx) used to authenticate</param>
        /// <param name="certificatePassword">Password used for the private key certificate (.pfx)</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <param name="loginName">
        /// Name of the user (login name) on whose behalf to create the access token.
        /// Supported input formats are SID and User Principal Name (UPN).
        /// If the parameter is left empty (including null) an App Only Context will be created.
        /// </param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppAuthenticatedContext(string siteUrl, string clientId, string certificatePath, SecureString certificatePassword, string certificateIssuerId, string loginName)
        {
            var certificate = new X509Certificate2(certificatePath, certificatePassword);
            return GetHighTrustCertificateAppAuthenticatedContext(siteUrl, clientId, certificate, certificateIssuerId, loginName);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="storeName">The name of the store for the certificate</param>
        /// <param name="storeLocation">The location of the store for the certificate</param>
        /// <param name="thumbPrint">The thumbprint of the certificate to locate in the store</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <param name="loginName">
        /// Name of the user (login name) on whose behalf to create the access token.
        /// Supported input formats are SID and User Principal Name (UPN).
        /// If the parameter is left empty (including null) an App Only Context will be created.
        /// </param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppAuthenticatedContext(string siteUrl, string clientId, StoreName storeName, StoreLocation storeLocation, string thumbPrint, string certificateIssuerId, string loginName)
        {
            // Retrieve the certificate from the Windows Certificate Store
            var cert = Utilities.X509CertificateUtility.LoadCertificate(storeName, storeLocation, thumbPrint);
            return GetHighTrustCertificateAppAuthenticatedContext(siteUrl, clientId, cert, certificateIssuerId, loginName);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using High Trust Certificate App Only Authentication
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The SharePoint Client ID</param>
        /// <param name="certificate">Private key certificate (.pfx) used to authenticate</param>
        /// <param name="certificateIssuerId">The IssuerID under which the CER counterpart of the PFX has been registered in SharePoint as a Trusted Security Token issuer</param>
        /// <param name="loginName">
        /// Name of the user (login name) on whose behalf to create the access token.
        /// Supported input formats are SID and User Principal Name (UPN).
        /// If the parameter is left empty (including null) an App Only Context will be created.
        /// </param>
        /// <returns>Authenticated SharePoint ClientContext</returns>
        public ClientContext GetHighTrustCertificateAppAuthenticatedContext(string siteUrl, string clientId, X509Certificate2 certificate, string certificateIssuerId, string loginName)
        {
            var siteUri = new Uri(siteUrl);
            var clientContext = new ClientContext(siteUri);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif

            // Feed the TokenHelper the SharePoint information so it doesn't try to fetch it from the config file
            TokenHelper.Realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            TokenHelper.ClientId = clientId;
            TokenHelper.ClientCertificate = certificate;
            TokenHelper.IssuerId = certificateIssuerId;

            // Configure the handler to generate the Bearer Access Token on each request and add it to the request
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                var accessToken = TokenHelper.GetS2SAccessTokenWithUserName(siteUri, loginName);
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return clientContext;
        }
#endif

        #endregion

        #region Authenticating against SharePoint Online using Azure AD based authentication
#if !ONPREMISES && !NETSTANDARD2_0

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app being registered in your Azure AD.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="userPrincipalName">The user id</param>
        /// <param name="userPassword">The user's password as a secure string</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADCredentialsContext(string siteUrl, string userPrincipalName, SecureString userPassword, AzureEnvironment environment = AzureEnvironment.Production)
        {
            string password = new System.Net.NetworkCredential(string.Empty, userPassword).Password;
            return GetAzureADCredentialsContext(siteUrl, userPrincipalName, password, environment);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app being registered in your Azure AD.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="userPrincipalName">The user id</param>
        /// <param name="userPassword">The user's password as a string</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADCredentialsContext(string siteUrl, string userPrincipalName, string userPassword, AzureEnvironment environment = AzureEnvironment.Production)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_GetContext, siteUrl);
            Log.Debug(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_TenantUser, userPrincipalName);

            var spUri = new Uri(siteUrl);
            string resourceUri = spUri.Scheme + "://" + spUri.Authority;

            var clientContext = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                EnsureAzureADCredentialsToken(resourceUri, userPrincipalName, userPassword, environment);
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + azureADCredentialsToken;
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.AzureADCredentials,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                UserName = userPrincipalName,
                Password = userPassword
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

        /// <summary>
        /// Acquires an access token using Azure AD credential flow. This depends on the SPO Management Shell app being registered in your Azure AD.
        /// </summary>
        /// <param name="resourceUri">Resouce to request access for</param>
        /// <param name="username">User id</param>
        /// <param name="password">Password</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Acces token</returns>
        public static async Task<string> AcquireTokenAsync(string resourceUri, string username, string password, AzureEnvironment environment)
        {
            HttpClient client = new HttpClient();
            string tokenEndpoint = $"{new AuthenticationManager().GetAzureADLoginEndPoint(environment)}/common/oauth2/token";

            var body = $"resource={resourceUri}&client_id=9bc3ab49-b65d-410a-85ad-de819febfddc&grant_type=password&username={HttpUtility.UrlEncode(username)}&password={HttpUtility.UrlEncode(password)}";
            var stringContent = new StringContent(body, System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");

            var result = await client.PostAsync(tokenEndpoint, stringContent).ContinueWith<string>((response) =>
            {
                return response.Result.Content.ReadAsStringAsync().Result;
            });

            JObject jobject = JObject.Parse(result);
            var token = jobject["access_token"].Value<string>();
            return token;
        }

        private void EnsureAzureADCredentialsToken(string resourceUri, string userPrincipalName, string userPassword, AzureEnvironment environment)
        {
            if (azureADCredentialsToken == null)
            {
                lock (tokenLock)
                {
                    if (azureADCredentialsToken == null)
                    {

                        String accessToken = Task.Run(() => AcquireTokenAsync(resourceUri, userPrincipalName, userPassword, environment)).GetAwaiter().GetResult();
                        ThreadPool.QueueUserWorkItem(obj =>
                        {
                            try
                            {
                                var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(accessToken);
                                Log.Debug(Constants.LOGGING_SOURCE, "Lease expiration date: {0}", token.ValidTo);
                                var lease = GetAccessTokenLease(token.ValidTo);
                                lease =
                                    TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
                                Thread.Sleep(lease);
                                azureADCredentialsToken = null;
                            }
                            catch (Exception ex)
                            {
                                Log.Warning(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManger_ProblemDeterminingTokenLease, ex);
                                azureADCredentialsToken = null;
                            }
                        });
                        azureADCredentialsToken = accessToken;
                    }
                }
            }
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Native Application Client ID</param>
        /// <param name="redirectUrl">The Azure AD Native Application Redirect Uri as a string</param>
        /// <param name="tokenCache">Optional token cache. If not specified an in-memory token cache will be used</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADNativeApplicationAuthenticatedContext(string siteUrl, string clientId, string redirectUrl, TokenCache tokenCache = null, AzureEnvironment environment = AzureEnvironment.Production)
        {
            return GetAzureADNativeApplicationAuthenticatedContext(siteUrl, clientId, new Uri(redirectUrl), tokenCache, environment);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Native Application Client ID</param>
        /// <param name="redirectUri">The Azure AD Native Application Redirect Uri</param>
        /// <param name="tokenCache">Optional token cache. If not specified an in-memory token cache will be used</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADNativeApplicationAuthenticatedContext(string siteUrl, string clientId, Uri redirectUri, TokenCache tokenCache = null, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var clientContext = new ClientContext(siteUrl);
            _contextUrl = siteUrl;
            _tokenCache = tokenCache;
            _clientId = clientId;
            _redirectUri = redirectUri;
            _commonAuthority = String.Format("{0}/common", GetAzureADLoginEndPoint(environment));

            clientContext.ExecutingWebRequest += clientContext_NativeApplicationExecutingWebRequest;

            return clientContext;
        }
#endif

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Web Application registered. The user will not be prompted for authentication, the current user's authentication context will be used by leveraging ADAL.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="accessTokenGetter">The AccessToken getter method to use</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADWebApplicationAuthenticatedContext(String siteUrl, Func<String, String> accessTokenGetter)
        {
            var clientContext = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                Uri resourceUri = new Uri(siteUrl);
                resourceUri = new Uri(resourceUri.Scheme + "://" + resourceUri.Host + "/");

                String accessToken = accessTokenGetter(resourceUri.ToString());
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return clientContext;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Web Application registered. The user will not be prompted for authentication, the current user's authentication context will be used by leveraging an explicit OAuth 2.0 Access Token value.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="accessToken">An explicit value for the AccessToken</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADAccessTokenAuthenticatedContext(String siteUrl, String accessToken)
        {
            var clientContext = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return clientContext;
        }

#if !NETSTANDARD2_0
        async void clientContext_NativeApplicationExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            var host = new Uri(_contextUrl);
            var ar = await AcquireNativeApplicationTokenAsync(_commonAuthority, host.Scheme + "://" + host.Host + "/");

            if (ar != null)
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
            }
        }
#endif

#if !NETSTANDARD2_0
        private async Task<AuthenticationResult> AcquireNativeApplicationTokenAsync(string authContextUrl, string resourceId)
        {
            AuthenticationResult ar = null;

            await new SynchronizationContextRemover();

            try
            {
                if (_tokenCache != null)
                {
                    _authContext = new AuthenticationContext(authContextUrl, _tokenCache);
                }
                else
                {

                    _authContext = new AuthenticationContext(authContextUrl);
                }

                if (_authContext.TokenCache.ReadItems().Any())
                {
                    string cachedAuthority =
                        _authContext.TokenCache.ReadItems().First().Authority;

                    if (_tokenCache != null)
                    {
                        _authContext = new AuthenticationContext(cachedAuthority, _tokenCache);
                    }
                    else
                    {
                        _authContext = new AuthenticationContext(cachedAuthority);
                    }
                }
                ar = (await _authContext.AcquireTokenSilentAsync(resourceId, _clientId));
            }
            catch (Exception)
            {
                //not in cache; we'll get it with the full oauth flow
            }

            if (ar == null)
            {
                try
                {
                    ar = await _authContext.AcquireTokenAsync(resourceId, _clientId, _redirectUri, new PlatformParameters(PromptBehavior.Always));

                }
                catch (Exception acquireEx)
                {
                    Log.Error(Constants.LOGGING_SOURCE, acquireEx.ToDetailedString());
                }
            }

            return ar;
        }
#endif

#if !NETSTANDARD2_0
        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="storeName">The name of the store for the certificate</param>
        /// <param name="storeLocation">The location of the store for the certificate</param>
        /// <param name="thumbPrint">The thumbprint of the certificate to locate in the store</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>ClientContext being used</returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, StoreName storeName, StoreLocation storeLocation, string thumbPrint, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var cert = Utilities.X509CertificateUtility.LoadCertificate(storeName, storeLocation, thumbPrint);

            return GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, cert, environment);
        }
#endif

#if !NETSTANDARD2_0

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="certificatePath">The path to the certificate (*.pfx) file on the file system</param>
        /// <param name="certificatePassword">Password to the certificate</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, string certificatePath, string certificatePassword, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var certPassword = Utilities.EncryptionUtility.ToSecureString(certificatePassword);

            return GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, certificatePath, certPassword, environment);
        }
#endif

#if !NETSTANDARD2_0

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="certificatePath">The path to the certificate (*.pfx) file on the file system</param>
        /// <param name="certificatePassword">Password to the certificate</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, string certificatePath, SecureString certificatePassword, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var certfile = System.IO.File.OpenRead(certificatePath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            var cert = new X509Certificate2(
                certificateBytes,
                certificatePassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);

            return GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, cert, environment);
        }
#endif

#if !NETSTANDARD2_0
        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="certificate">Certificate used to authenticate</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns></returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, X509Certificate2 certificate, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var clientContext = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif

            string authority = string.Format(CultureInfo.InvariantCulture, "{0}/{1}/", GetAzureADLoginEndPoint(environment), tenant);

            var authContext = new AuthenticationContext(authority);

            var clientAssertionCertificate = new ClientAssertionCertificate(clientId, certificate);

            var host = new Uri(siteUrl);

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                var ar = Task.Run(() => authContext
                    .AcquireTokenAsync(host.Scheme + "://" + host.Host + "/", clientAssertionCertificate))
                    .GetAwaiter().GetResult();
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.AzureADCertificate,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                ClientId = clientId,
                Tenant = tenant,
                Certificate = certificate,
                Environment = environment
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }
#endif

        /// <summary>
        /// Get's the Azure AD login end point for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure AD login endpoint</returns>
        public string GetAzureADLoginEndPoint(AzureEnvironment environment)
        {
            switch (environment)
            {
                case AzureEnvironment.Production:
                    {
                        return "https://login.microsoftonline.com";
                    }
                case AzureEnvironment.Germany:
                    {
                        return "https://login.microsoftonline.de";
                    }
                case AzureEnvironment.China:
                    {
                        return "https://login.chinacloudapi.cn";
                    }
                case AzureEnvironment.USGovernment:
                    {
                        return "https://login.microsoftonline.us";
                    }
                case AzureEnvironment.PPE:
                    {
                        return "https://login.windows-ppe.net";
                    }
                default:
                    {
                        return "https://login.microsoftonline.com";
                    }
            }
        }
        #endregion

#if !NETSTANDARD2_0
        #region Authenticating against SharePoint on-premises using ADFS based authentication
        /// <summary>
        /// Returns a SharePoint on-premises ClientContext for sites secured via ADFS
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="user">Name of the user (e.g. administrator) </param>
        /// <param name="password">Password of the user</param>
        /// <param name="domain">Windows domain of the user</param>
        /// <param name="sts">Hostname of the ADFS server (e.g. sts.company.com)</param>
        /// <param name="idpId">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetADFSUserNameMixedAuthenticatedContext(string siteUrl, string user, string password, string domain, string sts, string idpId, int logonTokenCacheExpirationWindow = 10)
        {

            ClientContext clientContext = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif
            clientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
            {
                if (fedAuth != null)
                {
                    Cookie fedAuthCookie = fedAuth.GetCookies(new Uri(siteUrl))["FedAuth"];
                    // If cookie is expired a new fedAuth cookie needs to be requested
                    if (fedAuthCookie == null || fedAuthCookie.Expires < DateTime.Now)
                    {
                        fedAuth = new UsernameMixed().GetFedAuthCookie(siteUrl, $"{domain}\\{user}", password, new Uri($"https://{sts}/adfs/services/trust/13/usernamemixed"), idpId, logonTokenCacheExpirationWindow);
                    }
                }
                else
                {
                    fedAuth = new UsernameMixed().GetFedAuthCookie(siteUrl, $"{domain}\\{user}", password, new Uri($"https://{sts}/adfs/services/trust/13/usernamemixed"), idpId, logonTokenCacheExpirationWindow);
                }

                if (fedAuth == null)
                {
                    throw new Exception("No fedAuth cookie acquired");
                }

                webRequestEventArgs.WebRequestExecutor.WebRequest.CookieContainer = fedAuth;
            };

            return clientContext;
        }

        /// <summary>
        /// Refreshes the SharePoint FedAuth cookie
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="user">Name of the user (e.g. administrator) </param>
        /// <param name="password">Password of the user</param>
        /// <param name="domain">Windows domain of the user</param>
        /// <param name="sts">Hostname of the ADFS server (e.g. sts.company.com)</param>
        /// <param name="idpId">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.</param>
        public void RefreshADFSUserNameMixedAuthenticatedContext(string siteUrl, string user, string password, string domain, string sts, string idpId, int logonTokenCacheExpirationWindow = 10)
        {
            fedAuth = new UsernameMixed().GetFedAuthCookie(siteUrl, $"{domain}\\{user}", password, new Uri($"https://{sts}/adfs/services/trust/13/usernamemixed"), idpId, logonTokenCacheExpirationWindow);
        }

        /// <summary>
        /// Returns a SharePoint on-premises ClientContext for sites secured via ADFS
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="serialNumber">Represents the serial number of the certificate as displayed by the certificate dialog box, but without the spaces, or as returned by the System.Security.Cryptography.X509Certificates.X509Certificate.GetSerialNumberString method</param>
        /// <param name="sts">Hostname of the ADFS server (e.g. sts.company.com)</param>
        /// <param name="idpId">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetADFSCertificateMixedAuthenticationContext(string siteUrl, string serialNumber, string sts, string idpId, int logonTokenCacheExpirationWindow = 10)
        {
            ClientContext clientContext = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif
            clientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
            {
                if (fedAuth != null)
                {
                    Cookie fedAuthCookie = fedAuth.GetCookies(new Uri(siteUrl))["FedAuth"];
                    // If cookie is expired a new fedAuth cookie needs to be requested
                    if (fedAuthCookie == null || fedAuthCookie.Expires < DateTime.Now)
                    {
                        fedAuth = new CertificateMixed().GetFedAuthCookie(siteUrl, serialNumber, new Uri($"https://{sts}/adfs/services/trust/13/certificatemixed"), idpId, logonTokenCacheExpirationWindow);
                    }
                }
                else
                {
                    fedAuth = new CertificateMixed().GetFedAuthCookie(siteUrl, serialNumber, new Uri($"https://{sts}/adfs/services/trust/13/certificatemixed"), idpId, logonTokenCacheExpirationWindow);
                }

                if (fedAuth == null)
                {
                    throw new Exception("No fedAuth cookie acquired");
                }

                webRequestEventArgs.WebRequestExecutor.WebRequest.CookieContainer = fedAuth;
            };
            return clientContext;
        }

        /// <summary>
        /// Refreshes the SharePoint FedAuth cookie
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="serialNumber">Represents the serial number of the certificate as displayed by the certificate dialog box, but without the spaces, or as returned by the System.Security.Cryptography.X509Certificates.X509Certificate.GetSerialNumberString method</param>
        /// <param name="sts">Hostname of the ADFS server (e.g. sts.company.com)</param>
        /// <param name="idpId">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public void RefreshADFSCertificateMixedAuthenticationContext(string siteUrl, string serialNumber, string sts, string idpId, int logonTokenCacheExpirationWindow = 10)
        {
            fedAuth = new CertificateMixed().GetFedAuthCookie(siteUrl, serialNumber, new Uri($"https://{sts}/adfs/services/trust/13/certificatemixed"), idpId, logonTokenCacheExpirationWindow);

        }

        /// <summary>
        /// Returns a SharePoint on-premises ClientContext for sites secured via ADFS
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="sts">Hostname of the ADFS server (e.g. sts.company.com)</param>
        /// <param name="idpId">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetADFSKerberosMixedAuthenticationContext(string siteUrl, string sts, string idpId, int logonTokenCacheExpirationWindow = 10)
        {
            ClientContext clientContext = new ClientContext(siteUrl);
#if !ONPREMISES || SP2016 || SP2019
            clientContext.DisableReturnValueCache = true;
#endif
            clientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
            {
                if (fedAuth != null)
                {
                    Cookie fedAuthCookie = fedAuth.GetCookies(new Uri(siteUrl))["FedAuth"];
                    // If cookie is expired a new fedAuth cookie needs to be requested
                    if (fedAuthCookie == null || fedAuthCookie.Expires < DateTime.Now)
                    {
                        fedAuth = new KerberosMixed().GetFedAuthCookie(siteUrl,
                            new Uri($"https://{sts}/adfs/services/trust/13/kerberosmixed"),
                            idpId,
                            logonTokenCacheExpirationWindow);
                    }
                }
                else
                {
                    fedAuth = new KerberosMixed().GetFedAuthCookie(siteUrl,
                        new Uri($"https://{sts}/adfs/services/trust/13/kerberosmixed"),
                        idpId,
                        logonTokenCacheExpirationWindow);
                }

                if (fedAuth == null)
                {
                    throw new Exception("No fedAuth cookie acquired");
                }

                webRequestEventArgs.WebRequestExecutor.WebRequest.CookieContainer = fedAuth;
            };
            return clientContext;
        }

        /// <summary>
        /// Refreshes the SharePoint FedAuth cookie
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="serialNumber">Certificate's serial number. Can be found in Serial number field in the certificate.</param>
        /// <param name="sts">Hostname of the ADFS server (e.g. sts.company.com)</param>
        /// <param name="idpId">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Optioanlly provide the value of the SharePoint STS logonTokenCacheExpirationWindow. Defaults to 10 minutes.</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public void RefreshADFSKerberosMixedAuthenticationContext(string siteUrl, string serialNumber, string sts, string idpId, int logonTokenCacheExpirationWindow = 10)
        {
            fedAuth = new KerberosMixed().GetFedAuthCookie(siteUrl,
                new Uri($"https://{sts}/adfs/services/trust/13/kerberosmixed"),
                idpId,
                logonTokenCacheExpirationWindow);
        }

        public static void GetAdfsConfigurationFromTargetUri(Uri targetApplicationUri, string loginProviderName, out string adfsHost, out string adfsRelyingParty)
        {
            adfsHost = "";
            adfsRelyingParty = "";

            var trustEndpoint = new Uri(new Uri(targetApplicationUri.GetLeftPart(UriPartial.Authority)), !string.IsNullOrWhiteSpace(loginProviderName) ? $"/_trust/?trust={loginProviderName}" : "/_trust/");
            var request = (HttpWebRequest)WebRequest.Create(trustEndpoint);
            request.AllowAutoRedirect = false;

            try
            {
                using (var response = request.GetResponse())
                {
                    var locationHeader = response.Headers["Location"];
                    if (locationHeader != null)
                    {
                        var redirectUri = new Uri(locationHeader);
                        Dictionary<string, string> queryParameters = Regex.Matches(redirectUri.Query, "([^?=&]+)(=([^&]*))?").Cast<Match>().ToDictionary(x => x.Groups[1].Value, x => Uri.UnescapeDataString(x.Groups[3].Value));
                        adfsHost = redirectUri.Host;
                        adfsRelyingParty = queryParameters["wtrealm"];
                    }
                }
            }
            catch (WebException ex)
            {
                throw new Exception("Endpoint does not use ADFS for authentication.", ex);
            }
        }

        #endregion
#endif
    }
}
