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
    public class WindowsTransport : BaseProvider
    {
        /// <summary>
        /// Performs active authentication against ADFS using the trust/13/usernamemixed ADFS endpoint.
        /// </summary>
        /// <param name="siteUrl">Url of the SharePoint site that's secured via ADFS</param>
        /// <param name="credentials">Credentials to authenticate against endpoint.</param>
        /// <param name="windowsTransport">Uri to the ADFS usernamemixed endpoint</param>
        /// <param name="relyingPartyIdentifier">Identifier of the ADFS relying party that we're hitting</param>
        /// <param name="logonTokenCacheExpirationWindow"></param>
        /// <returns>A cookiecontainer holding the FedAuth cookie</returns>
        public CookieContainer GetFedAuthCookie(string siteUrl, NetworkCredential credentials, Uri windowsTransport, string relyingPartyIdentifier, int logonTokenCacheExpirationWindow)
        {
            WindowsTransport adfsTokenProvider = new WindowsTransport();
            var token = adfsTokenProvider.RequestToken(credentials, windowsTransport, relyingPartyIdentifier);
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

        private GenericXmlSecurityToken RequestToken(NetworkCredential credentials, Uri windowsTransport, string relyingPartyIdentifier)
        {
            GenericXmlSecurityToken genericToken = null;
            using (var factory = new WSTrustChannelFactory(new WindowsWSTrustBinding(), new EndpointAddress(windowsTransport)))
            {
                factory.TrustVersion = TrustVersion.WSTrust13;
                // Hookup the credentials
                factory.Credentials.Windows.ClientCredential = credentials;

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