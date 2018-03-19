using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class WebNameToken : TokenDefinition
    {
        public WebNameToken(Web web) : base(web, "{webname}")
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