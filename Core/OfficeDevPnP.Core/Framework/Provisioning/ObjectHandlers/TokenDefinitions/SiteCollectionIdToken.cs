using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionIdToken : TokenDefinition
    {
        public SiteCollectionIdToken(Web web)
            : base(web, "~sitecollectionid", "{sitecollectionid}")
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
                    CacheValue = context.Site.Id.ToString();
                }
            }
            return CacheValue;
        }
    }
}