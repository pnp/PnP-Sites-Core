using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteToken : TokenDefinition
    {
        public SiteToken(Web web)
            : base(web, "~site", "{site}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    context.Load(context.Web, w => w.ServerRelativeUrl);
                    context.ExecuteQueryRetry();
                    CacheValue = context.Web.ServerRelativeUrl.TrimEnd('/');
                }
            }
            return CacheValue;
        }
    }
}