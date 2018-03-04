/* Based on reflectored code coming from Microsoft.IdentityModel.Protocols.WSTrust.Bindings.WSTrustBindingBase class */
#if !NETSTANDARD2_0
using System;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Security;
using System.ServiceModel.Security.Tokens;

namespace OfficeDevPnP.Core.IdentityModel.WSTrustBindings
{
    /// <summary>
    /// Class holds methos to enable applications to construct trusted message exchanges
    /// </summary>
    public abstract class WSTrustBinding : Binding
    {
        private bool _enableRsaProofKeys;
        private SecurityMode _securityMode;
        private TrustVersion _trustVersion;

        protected abstract void ApplyTransportSecurity(HttpTransportBindingElement transport);
        protected abstract SecurityBindingElement CreateSecurityBindingElement();

        protected WSTrustBinding(SecurityMode securityMode) : this(securityMode, TrustVersion.WSTrust13)
        { 
        }

        protected WSTrustBinding(SecurityMode securityMode, TrustVersion trustVersion)
        {
            _securityMode = SecurityMode.Message;
            _trustVersion = TrustVersion.WSTrust13;
          
            if (trustVersion == null)
            {
                throw new ArgumentNullException(nameof(trustVersion));
            }
            
            ValidateTrustVersion(trustVersion);
            ValidateSecurityMode(securityMode);
            _securityMode = securityMode;
            _trustVersion = trustVersion;
        }

        protected virtual SecurityBindingElement ApplyMessageSecurity(SecurityBindingElement securityBindingElement)
        {
            if (securityBindingElement == null)
            {
                throw new ArgumentNullException(nameof(securityBindingElement));
            }
            
            if (TrustVersion.WSTrustFeb2005 == _trustVersion)
            {
                securityBindingElement.MessageSecurityVersion = MessageSecurityVersion.WSSecurity11WSTrustFebruary2005WSSecureConversationFebruary2005WSSecurityPolicy11BasicSecurityProfile10;
            }
            else
            {
                securityBindingElement.MessageSecurityVersion = MessageSecurityVersion.WSSecurity11WSTrust13WSSecureConversation13WSSecurityPolicy12BasicSecurityProfile10;
            }
            
            if (_enableRsaProofKeys)
            {
                RsaSecurityTokenParameters item = new RsaSecurityTokenParameters
                {
                    InclusionMode = SecurityTokenInclusionMode.Never,
                    RequireDerivedKeys = false
                };
                securityBindingElement.OptionalEndpointSupportingTokenParameters.Endorsing.Add(item);
            }
            
            return securityBindingElement;
        }

        /// <summary>
        /// Creates binding elements required for trust message exchange
        /// </summary>
        /// <returns>Returns Binding elements created</returns>
        public override BindingElementCollection CreateBindingElements()
        {
            BindingElementCollection elements = new BindingElementCollection();
            elements.Clear();
            if ((SecurityMode.Message == _securityMode) || (SecurityMode.TransportWithMessageCredential == _securityMode))
            {
                elements.Add(ApplyMessageSecurity(CreateSecurityBindingElement()));
            }
            elements.Add(CreateEncodingBindingElement());
            elements.Add(CreateTransportBindingElement());
            return elements.Clone();
        }

        protected virtual MessageEncodingBindingElement CreateEncodingBindingElement()
        {
            return new TextMessageEncodingBindingElement { ReaderQuotas = { MaxArrayLength = 0x200000, MaxStringContentLength = 0x200000 } };
        }

        protected virtual HttpTransportBindingElement CreateTransportBindingElement()
        {
            HttpTransportBindingElement element;
            
            if (SecurityMode.Message == _securityMode)
            {
                element = new HttpTransportBindingElement();
            }
            else
            {
                element = new HttpsTransportBindingElement();
            }
            
            element.MaxReceivedMessageSize = 0x200000L;
            
            if (SecurityMode.Transport == _securityMode)
            {
                ApplyTransportSecurity(element);
            }
            
            return element;
        }

        protected static void ValidateSecurityMode(SecurityMode securityMode)
        {
            if (((securityMode != SecurityMode.None) && (securityMode != SecurityMode.Message)) && ((securityMode != SecurityMode.Transport) && (securityMode != SecurityMode.TransportWithMessageCredential)))
            {
                throw new ArgumentOutOfRangeException(nameof(securityMode));
            }
            
            if (securityMode == SecurityMode.None)
            {
                throw new InvalidOperationException("ID3224");
            }
        }

        protected void ValidateTrustVersion(TrustVersion trustVersion)
        {
            if ((trustVersion != TrustVersion.WSTrust13) && (trustVersion != TrustVersion.WSTrustFeb2005))
            {
                throw new ArgumentOutOfRangeException(nameof(trustVersion));
            }
        }

        /// <summary>
        /// Based on this values enables RSA key to encrypt the message
        /// </summary>
        public bool EnableRsaProofKeys
        {
            get
            {
                return _enableRsaProofKeys;
            }
            set
            {
                _enableRsaProofKeys = value;
            }
        }

        /// <summary>
        /// Gets the encryption scheme
        /// </summary>
        public override string Scheme
        {
            get
            {
                TransportBindingElement element = CreateBindingElements().Find<TransportBindingElement>();

                if (element == null)
                {
                    return string.Empty;
                }
                
                return element.Scheme;
            }
        }

        /// <summary>
        /// Gets or sets Security mode
        /// </summary>
        public SecurityMode SecurityMode
        {
            get
            {
                return _securityMode;
            }
            set
            {
                ValidateSecurityMode(value);
                _securityMode = value;
            }
        }

        /// <summary>
        /// Gets or sets the version
        /// </summary>
        public TrustVersion TrustVersion
        {
            get
            {
                return _trustVersion;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }
                
                ValidateTrustVersion(value);
                _trustVersion = value;
            }
        }
    }
}
#endif