using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteIdEncodedToken : TokenDefinition
    {
        public SiteIdEncodedToken(Web web)
            : base(web, "~siteidencoded", "{siteidencoded}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    context.Load(context.Web, w => w.Id);
                    context.ExecuteQueryRetry();
                    CacheValue = context.Web.Id.ToString().Replace("-", "%2D");
                }
            }
            return CacheValue;
        }
    }
}