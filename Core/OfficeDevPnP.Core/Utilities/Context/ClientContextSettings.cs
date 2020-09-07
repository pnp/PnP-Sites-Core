using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Security.Cryptography.X509Certificates;

namespace OfficeDevPnP.Core.Utilities.Context
{
    internal class ClientContextSettings
    {
        #region properties
        // Generic
        internal ClientContextType Type { get; set; }
        internal string SiteUrl { get; set; }
        internal AuthenticationManager AuthenticationManager { get; set; }

        // User name + password flows
        internal string UserName { get; set; }
        internal string Password { get; set; }
        
        // App Only flows
        internal string ClientId { get; set; }
        internal string ClientSecret { get; set; }
        internal string Realm { get; set; }
        internal string AcsHostUrl { get; set; }
        internal string GlobalEndPointPrefix { get; set; }
        internal string Tenant { get; set; }
        internal X509Certificate2 Certificate { get; set; }

        internal IClientAssertionCertificate ClientAssertionCertificate { get; set; }

        internal AzureEnvironment Environment { get; set; }
        #endregion

        #region methods
        internal bool UsesDifferentAudience(string newSiteUrl)
        {
            Uri newAudience = new Uri(newSiteUrl);
            Uri currentAudience = new Uri(this.SiteUrl); 

            if (newAudience.Host != currentAudience.Host)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion


    }
}
