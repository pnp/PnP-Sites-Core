using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{siteowner}",
       Description = "Returns the login name of the current site owner",
       Example = "{siteowner}",
       Returns = "i:0#.f|membership|user@domain.com")]
    internal class SiteOwnerToken : TokenDefinition
    {
        public SiteOwnerToken(Web web)
            : base(web, "~siteowner", "{siteowner}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                var site = TokenContext.Site;
                TokenContext.Load(site, s => s.Owner);
                TokenContext.ExecuteQueryRetry();
                CacheValue = site.Owner.LoginName;
            }
            return CacheValue;
        }
    }
}