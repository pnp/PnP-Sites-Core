using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionIdEncodedToken : TokenDefinition
    {
        public SiteCollectionIdEncodedToken(Web web)
            : base(web, "~sitecollectionidencoded", "{sitecollectionidencoded}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    context.Load(context.Site, s => s.Id);
                    context.ExecuteQueryRetry();
                    CacheValue = context.Site.Id.ToString().Replace("-", "%2D");
                }
            }
            return CacheValue;
        }
    }
}