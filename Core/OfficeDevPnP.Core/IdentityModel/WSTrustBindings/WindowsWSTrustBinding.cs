/* Based on reflectored code coming from Microsoft.IdentityModel.Protocols.WSTrust.Bindings.WindowsWSTrustBinding class */
#if !NETSTANDARD2_0
using System;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace OfficeDevPnP.Core.IdentityModel.WSTrustBindings
{
    public class WindowsWSTrustBinding : WSTrustBinding
    {
        public WindowsWSTrustBinding()
          : this(SecurityMode.Message)
        {
        }

        public WindowsWSTrustBinding(SecurityMode securityMode)
          : base(securityMode)
        {
        }

        protected override SecurityBindingElement CreateSecurityBindingElement()
        {
            if (SecurityMode.Message == this.SecurityMode)
                return (SecurityBindingElement)SecurityBindingElement.CreateSspiNegotiationBindingElement(true);
            if (SecurityMode.TransportWithMessageCredential == this.SecurityMode)
                return (SecurityBindingElement)SecurityBindingElement.CreateSspiNegotiationOverTransportBindingElement(true);
            return (SecurityBindingElement)null;
        }

        protected override void ApplyTransportSecurity(HttpTransportBindingElement transport)
        {
            transport.AuthenticationScheme = AuthenticationSchemes.Negotiate;
        }
    }
}
#endif