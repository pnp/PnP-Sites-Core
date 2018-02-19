using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class AuthenticationRealmToken : TokenDefinition
    {
        public AuthenticationRealmToken(Web web)
            : base(web, "~authenticationrealm", "~realm", "{authenticationrealm}", "{realm}")
        {
        }
        public override string GetReplaceValue()
        {
#if !NETSTANDARD2_0
            return Web.GetAuthenticationRealm().ToString();
#else
            throw new Exception("authenticationrealm token not supported");
#endif
        }
    }
}
