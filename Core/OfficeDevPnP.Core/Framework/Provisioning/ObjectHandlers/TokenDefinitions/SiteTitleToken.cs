using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteTitleToken : TokenDefinition
    {
        public SiteTitleToken(Web web) : base(web, "{sitetitle}", "~sitename", "{sitename}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    context.Load(context.Web, w => w.Title);
                    context.ExecuteQueryRetry();
                    CacheValue = context.Web.Title;
                }
            }
            return CacheValue;
        }
    }
}