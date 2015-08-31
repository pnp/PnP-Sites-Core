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
                var context = this.Web.Context as ClientContext;
                context.Load(Web, w => w.Id);
                context.ExecuteQueryRetry();
                CacheValue = Web.Id.ToString();
            }
            return CacheValue;
        }
    }
}