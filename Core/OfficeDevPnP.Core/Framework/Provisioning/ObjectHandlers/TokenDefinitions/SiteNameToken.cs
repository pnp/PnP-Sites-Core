using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteNameToken : TokenDefinition
    {
        public SiteNameToken(Web web)
            : base(web, "~sitename", "{sitename}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                CacheValue = this.Web.GetName();
            }
            return CacheValue;
        }
    }
}