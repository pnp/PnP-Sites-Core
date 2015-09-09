/* Based on reflectored code coming from Microsoft.IdentityModel.Protocols.WSTrust.Bindings.UserNameWSTrustBinding class */

using System;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace OfficeDevPnP.Core.IdentityModel.WSTrustBindings
{
    public class WindowsWSTrustBinding : WSTrustBinding
    {
        // Methods
        public WindowsWSTrustBinding() : base(SecurityMode.Transport) { }

        protected override void ApplyTransportSecurity(HttpTransportBindingElement transport) {
            transport.AuthenticationScheme = AuthenticationSchemes.Negotiate;
        }


        protected override SecurityBindingElement CreateSecurityBindingElement()
        {
            return SecurityBindingElement.CreateKerberosOverTransportBindingElement();
        }
   }
}
