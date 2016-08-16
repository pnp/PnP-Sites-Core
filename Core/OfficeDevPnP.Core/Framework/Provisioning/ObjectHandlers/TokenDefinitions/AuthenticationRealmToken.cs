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
            return Web.GetAuthenticationRealm().ToString();
        }
    }
}
