using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class EveryoneButExternalUsersToken : TokenDefinition
    {
        public EveryoneButExternalUsersToken(Web web)
            : base(web, $"{{everyonebutexternalusers}}")
        {

        }

        public override string GetReplaceValue()
        {
           return $"c:0-.f|rolemanager|spo-grid-all-users/{Web.GetAuthenticationRealm()}";
        }
    }
}
