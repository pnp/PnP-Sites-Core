#if !ONPREMISES
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{everyonebutexternalusers}",
       Description = "Returns the claim for everyone but external users in this tenant",
       Example = "{everyonebutexternalusers}",
       Returns = "c:0-.f|rolemanager|spo-grid-all-users/b6e37e85-1739-4512-888c-2078dc575169")]
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
#endif