/* Based on reflectored code coming from Microsoft.IdentityModel.Protocols.WSTrust.Bindings.UserNameWSTrustBinding class */
#if !NETSTANDARD2_0
using System;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace OfficeDevPnP.Core.IdentityModel.WSTrustBindings
{
    /// <summary>
    /// Class holds methods and properties for certificate trust binding
    /// </summary>
    public class CertificateWSTrustBinding : WSTrustBinding
    {
        // Fields
        private HttpClientCredentialType _clientCredentialType;

        // Methods
        /// <summary>
        /// Deafult Constructor
        /// </summary>
        public CertificateWSTrustBinding() : this(SecurityMode.Message, HttpClientCredentialType.None)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="securityMode"></param>
        public CertificateWSTrustBinding(SecurityMode securityMode) : base(securityMode)
        {
            if (SecurityMode.Message == securityMode)
            {
                _clientCredentialType = HttpClientCredentialType.None;
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="mode"></param>
        /// <param name="clientCredentialType"></param>
        public CertificateWSTrustBinding(SecurityMode mode, HttpClientCredentialType clientCredentialType) : base(mode)
        {
            if (!IsHttpClientCredentialTypeDefined(clientCredentialType))
            {
                throw new ArgumentOutOfRangeException(nameof(clientCredentialType));
            }

            if (((SecurityMode.Transport == mode) && (HttpClientCredentialType.Digest != clientCredentialType)) && (HttpClientCredentialType.Basic != clientCredentialType))
            {
                throw new InvalidOperationException("ID3225");
            }

            _clientCredentialType = clientCredentialType;
        }

        protected override void ApplyTransportSecurity(HttpTransportBindingElement transport)
        {
            if (_clientCredentialType == HttpClientCredentialType.Certificate)
            {
                transport.AuthenticationScheme = AuthenticationSchemes.Negotiate;
            }
            else
            {
                transport.AuthenticationScheme = AuthenticationSchemes.Negotiate;
            }
        }

        protected override SecurityBindingElement CreateSecurityBindingElement()
        {
            if (SecurityMode.Message == base.SecurityMode)
            {
                return SecurityBindingElement.CreateCertificateSignatureBindingElement();
            }

            if (SecurityMode.Transport == base.SecurityMode)
            {
                //return SecurityBindingElement.CreateUserNameOverTransportBindingElement();
                return SecurityBindingElement.CreateCertificateOverTransportBindingElement();
            }

            if (SecurityMode.TransportWithMessageCredential == base.SecurityMode)
            {
                return SecurityBindingElement.CreateCertificateOverTransportBindingElement();
            }

            return null;
        }

        private static bool IsHttpClientCredentialTypeDefined(HttpClientCredentialType value)
        {
            if ((((value != HttpClientCredentialType.None) && (value != HttpClientCredentialType.Basic)) && ((value != HttpClientCredentialType.Digest) && (value != HttpClientCredentialType.Ntlm))) && (value != HttpClientCredentialType.Windows))
            {
                return (value == HttpClientCredentialType.Certificate);
            }

            return true;
        }

        // Properties
        /// <summary>
        /// Gets or sets Http client credential type
        /// </summary>
        public HttpClientCredentialType ClientCredentialType
        {
            get
            {
                return _clientCredentialType;
            }
            set
            {
                if (!IsHttpClientCredentialTypeDefined(value))
                {
                    throw new ArgumentOutOfRangeException(nameof(value));
                }
                if (((SecurityMode.Transport == base.SecurityMode) && (HttpClientCredentialType.Digest != value)) && (HttpClientCredentialType.Basic != value))
                {
                    throw new InvalidOperationException("ID3225");
                }
                _clientCredentialType = value;
            }
        }
    }
}
#endif