/* Based on reflectored code coming from Microsoft.IdentityModel.Protocols.WSTrust.Bindings.KerberosWSTrustBinding class */
#if !NETSTANDARD2_0
using System;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace OfficeDevPnP.Core.IdentityModel.WSTrustBindings
{
    public class KerberosWSTrustBinding : WSTrustBinding
    {
        public KerberosWSTrustBinding()
          : this(SecurityMode.TransportWithMessageCredential)
        {
        }

        public KerberosWSTrustBinding(SecurityMode mode)
          : base(mode)
        {
        }

        protected override SecurityBindingElement CreateSecurityBindingElement()
        {
            if (SecurityMode.Message == this.SecurityMode)
                return (SecurityBindingElement)SecurityBindingElement.CreateKerberosBindingElement();
            if (SecurityMode.TransportWithMessageCredential == this.SecurityMode)
                return (SecurityBindingElement)SecurityBindingElement.CreateKerberosOverTransportBindingElement();
            return (SecurityBindingElement)null;
        }

        protected override void ApplyTransportSecurity(HttpTransportBindingElement transport)
        {
            transport.AuthenticationScheme = AuthenticationSchemes.Negotiate;
        }
    }
}
#endif