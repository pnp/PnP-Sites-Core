using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
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
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    var currentUser = context.Web.EnsureProperty(w => w.CurrentUser);
                    CacheValue = currentUser.Id.ToString();
                }
            }
            return CacheValue;
        }
    }
}