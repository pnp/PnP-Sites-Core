#if !NETSTANDARD2_0
using OfficeDevPnP.Core.IdentityModel.WSTrustBindings;
using System;
using System.IdentityModel.Protocols.WSTrust;
using System.IdentityModel.Tokens;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.ServiceModel.Security;

namespace OfficeDevPnP.Core.IdentityModel.TokenProviders.ADFS
{
    /// <summary>
    /// ADFS Active authentication based on username + password. Uses the trust/13/usernamemixed ADFS endpoint.
    /// </summary>
    public class CertificateMixed : BaseProvider
    {
        /// <summary>
        /// Performs active authentication against ADFS using the trust/13/usernamemixed ADFS endpoint.
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="serialNumber">Serial Number of the Current User > My Certificate to use to authenticate </param>
        /// <param name="certificateMixed">Uri to the ADFS certificatemixed endpoint</param>
        /// <param name="relyingPartyIdentifier">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Logon TokenCache expiration window integer value</param>
        /// <returns>A cookiecontainer holding the FedAuth cookie</returns>
        public CookieContainer GetFedAuthCookie(string siteUrl, string serialNumber, Uri certificateMixed, string relyingPartyIdentifier, int logonTokenCacheExpirationWindow)
        {
            CertificateMixed adfsTokenProvider = new CertificateMixed();

            var token = adfsTokenProvider.RequestToken(serialNumber, certificateMixed, relyingPartyIdentifier);
            string fedAuthValue = TransformSamlTokenToFedAuth(token.TokenXml.OuterXml, siteUrl, relyingPartyIdentifier);

            // Construct the cookie expiration date
            TimeSpan lifeTime = SamlTokenlifeTime(token.TokenXml.OuterXml);
            if (lifeTime == TimeSpan.Zero)
            {
                lifeTime = new TimeSpan(0, 60, 0);
            }

            int cookieLifeTime = Math.Min((lifeTime.Hours * 60 + lifeTime.Minutes), logonTokenCacheExpirationWindow);
            DateTime expiresOn = DateTime.Now.AddMinutes(cookieLifeTime);

            CookieContainer cc = null;

            if (!string.IsNullOrEmpty(fedAuthValue))
            {
                cc = new CookieContainer();
                Cookie samlAuth = new Cookie("FedAuth", fedAuthValue);
                samlAuth.Expires = expiresOn;
                samlAuth.Path = "/";
                samlAuth.Secure = true;
                samlAuth.HttpOnly = true;
                Uri samlUri = new Uri(siteUrl);
                samlAuth.Domain = samlUri.Host;
                cc.Add(samlAuth);
            }

            return cc;
        }

        /// <summary>
        /// Returns Generic XML Security Token from ADFS to generated FedAuth
        /// </summary>
        /// <param name="serialNumber">Serial Number of Certificate from CurrentUSer > My Certificate</param>
        /// <param name="certificateMixed">ADFS Endpoint for Certificate Mixed Authentication</param>
        /// <param name="relyingPartyIdentifier">Identifier of the ADFS relying party that we're hitting</param>
        /// <returns></returns>
        private GenericXmlSecurityToken RequestToken(string serialNumber, Uri certificateMixed, string relyingPartyIdentifier)
        {
            GenericXmlSecurityToken genericToken = null;
            using (var factory = new WSTrustChannelFactory(new CertificateWSTrustBinding(SecurityMode.TransportWithMessageCredential), new EndpointAddress(certificateMixed)))
            {
                
                factory.TrustVersion = TrustVersion.WSTrust13;
                // Hookup the user and password 
                factory.Credentials.ClientCertificate.SetCertificate(StoreLocation.CurrentUser, StoreName.My, X509FindType.FindBySerialNumber, serialNumber);
                
                var requestSecurityToken = new RequestSecurityToken
                {
                    RequestType = RequestTypes.Issue,
                    AppliesTo = new EndpointReference(relyingPartyIdentifier),
                    KeyType = KeyTypes.Bearer
                };

                IWSTrustChannelContract channel = factory.CreateChannel();
                genericToken = channel.Issue(requestSecurityToken) as GenericXmlSecurityToken;
                factory.Close();
            }
            return genericToken;
        }

    }
}
#endif