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
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    var currentUser = context.Web.EnsureProperty(w => w.CurrentUser);

                    CacheValue = currentUser.LoginName;
                }
            }
            return CacheValue;
        }
    }
}