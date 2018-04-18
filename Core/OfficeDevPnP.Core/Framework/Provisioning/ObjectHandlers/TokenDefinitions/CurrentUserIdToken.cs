using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
    Token = "{currentuserid}",
    Description = "Returns the ID of the current user e.g. the user using the engine.",
    Example = "{currentuserid}",
    Returns = "4")]
    internal class CurrentUserIdToken : TokenDefinition
    {
        public CurrentUserIdToken(Web web)
            : base(web, "~currentuserid", "{currentuserid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                var currentUser = TokenContext.Web.EnsureProperty(w => w.CurrentUser);
                CacheValue = currentUser.Id.ToString();
            }
            return CacheValue;
        }
    }
}