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
                var context = this.Web.Context as ClientContext;
                var currentUser = context.Web.EnsureProperty(w => w.CurrentUser);

                CacheValue = currentUser.Title;
            }
            return CacheValue;
        }
    }
}