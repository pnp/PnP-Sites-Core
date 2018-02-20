using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteNameToken : TokenDefinition
    {
        public SiteNameToken(Web web)
            //Due to backwardscompatibility issues this token can not use the intended sitename token
            //This is because SiteTitleToken historically was created with sitename token and incorrectly returned the site title.
            //If possible this should be changed to sitename in the future.
            //: base(web, "~sitename", "{sitename}")
            : base(web, "~webname", "{webname}")
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