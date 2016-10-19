using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionToken : TokenDefinition
    {
        public SiteCollectionToken(Web web)
            : base(web, "~sitecollection", "{sitecollection}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    var site = context.Site;
                    context.Load(site, s => s.RootWeb.ServerRelativeUrl);
                    context.ExecuteQueryRetry();
                    CacheValue = site.RootWeb.ServerRelativeUrl.TrimEnd('/');
                }
            }
            return CacheValue;
        }
    }
}