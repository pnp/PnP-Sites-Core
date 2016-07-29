using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class CurrentUserFullNameToken : TokenDefinition
    {
        public CurrentUserFullNameToken(Web web)
            : base(web, "~currentuserfullname", "{currentuserfullname}")
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
                    CacheValue = currentUser.Title;
                }
            }
            return CacheValue;
        }
    }
}