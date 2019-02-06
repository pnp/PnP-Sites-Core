#if !NETSTANDARD2_0
using OfficeDevPnP.Core.IdentityModel.WSTrustBindings;
using System;
using System.IdentityModel.Protocols.WSTrust;
using System.IdentityModel.Tokens;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Security;

namespace OfficeDevPnP.Core.IdentityModel.TokenProviders.ADFS
{
    /// <summary>
    /// ADFS Active authentication based on username + password. Uses the trust/13/usernamemixed ADFS endpoint.
    /// </summary>
    public class KerberosMixed : BaseProvider
    {
        /// <summary>
        /// Performs active authentication against ADFS using the trust/13/usernamemixed ADFS endpoint.
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="kerberosMixed">Uri to the ADFS kerberosMixed endpoint</param>
        /// <param name="relyingPartyIdentifier">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow">Logon TokenCache expiration window integer value</param>
        /// <returns>A cookiecontainer holding the FedAuth cookie</returns>
        public CookieContainer GetFedAuthCookie(string siteUrl, Uri kerberosMixed, string relyingPartyIdentifier, int logonTokenCacheExpirationWindow)
        {
            KerberosMixed adfsTokenProvider = new KerberosMixed();

            var token = adfsTokenProvider.RequestToken(kerberosMixed, relyingPartyIdentifier);
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
        /// <param name="kerberosMixed">ADFS Endpoint for Kerberos Mixed Authentication</param>
        /// <param name="relyingPartyIdentifier">Identifier of the ADFS relying party that we're hitting</param>
        /// <returns></returns>
        private GenericXmlSecurityToken RequestToken(Uri kerberosMixed, string relyingPartyIdentifier)
        {
            GenericXmlSecurityToken genericToken = null;
            using (var factory = new WSTrustChannelFactory(new KerberosWSTrustBinding(SecurityMode.TransportWithMessageCredential), new EndpointAddress(kerberosMixed)))
            {
                factory.TrustVersion = TrustVersion.WSTrust13;
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