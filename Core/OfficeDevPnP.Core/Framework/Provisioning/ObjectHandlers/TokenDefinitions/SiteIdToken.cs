using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteIdToken : TokenDefinition
    {
        public SiteIdToken(Web web)
            : base(web, "~siteid", "{siteid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                //var context = this.Web.Context as ClientContext;
                //context.Load(Web, w => w.Id);
                //context.ExecuteQueryRetry();
                //CacheValue = Web.Id.ToString();

                this.Web.EnsureProperty(w => w.Url);
                using (ClientContext context = this.Web.Context.Clone(this.Web.Url))
                {
                    context.Load(context.Web, w => w.Id);
                    context.ExecuteQueryRetry();
                    CacheValue = context.Web.Id.ToString();
                }
            }
            return CacheValue;
        }
    }
}