using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectionidencoded}",
        Description = "Returns the id of the site collection",
        Example = "{sitecollectionidencoded}",
        Returns = "767bc144-e605-4d8c-885a-3a980feb39c6")]
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
                TokenContext.Load(TokenContext.Site, s => s.Id);
                TokenContext.ExecuteQueryRetry();
                CacheValue = TokenContext.Site.Id.ToString();
            }
            return CacheValue;
        }
    }
}