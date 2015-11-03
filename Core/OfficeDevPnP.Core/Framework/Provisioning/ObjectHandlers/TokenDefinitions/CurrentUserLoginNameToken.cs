using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class CurrentUserLoginNameToken : TokenDefinition
    {
        public CurrentUserLoginNameToken(Web web)
            : base(web, "~currentuserloginname", "{currentuserloginname}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                var context = this.Web.Context as ClientContext;
                var currentUser = context.Web.EnsureProperty(w => w.CurrentUser);

                CacheValue = currentUser.LoginName;
            }
            return CacheValue;
        }
    }
}