#if !ONPREMISES
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
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
            string userIdentity = "";
            try
            {
                // New tenant
                userIdentity = $"c:0-.f|rolemanager|spo-grid-all-users/{this.TokenContext.Web.GetAuthenticationRealm()}";
                var spReader = this.TokenContext.Web.EnsureUser(userIdentity);
                this.TokenContext.Web.Context.Load(spReader);
                this.TokenContext.Web.Context.ExecuteQueryRetry();
            }
            catch (ServerException)
            {
                try
                {
                    // Old tenants
                    string claimName = this.TokenContext.Web.GetEveryoneExceptExternalUsersClaimName();
                    var claim = Utility.ResolvePrincipal(this.TokenContext, this.TokenContext.Web, claimName, PrincipalType.SecurityGroup, PrincipalSource.RoleProvider, null, false);
                    this.TokenContext.ExecuteQueryRetry();
                    userIdentity = claim.Value.LoginName;
                }
                catch { }
            }

            return userIdentity;
        }
    }
}
#endif