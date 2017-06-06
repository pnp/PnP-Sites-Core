/* Based on reflectored code coming from Microsoft.IdentityModel.Protocols.WSTrust.Bindings.UserNameWSTrustBinding class */

using System;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace OfficeDevPnP.Core.IdentityModel.WSTrustBindings
{
    public class CertificateWSTrustBinding : WSTrustBinding
    {
        // Fields
        private HttpClientCredentialType _clientCredentialType;

        // Methods
        public CertificateWSTrustBinding() : this(SecurityMode.Message, HttpClientCredentialType.None)
        {
        }

        public CertificateWSTrustBinding(SecurityMode securityMode) : base(securityMode)
        {
            if (SecurityMode.Message == securityMode)
            {
                _clientCredentialType = HttpClientCredentialType.None;
            }
        }

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

