using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectionidencoded}",
        Description = "Returns the HTML safe id of the site collection",
        Example = "{sitecollectionidencoded}",
        Returns = "767bc144%2De605%2D4d8c%2D885a%2D3a980feb39c6")]
    internal class SiteCollectionIdEncodedToken : TokenDefinition
    {
        public SiteCollectionIdEncodedToken(Web web)
            : base(web, "~sitecollectionidencoded", "{sitecollectionidencoded}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Load(TokenContext.Site, s => s.Id);
                TokenContext.ExecuteQueryRetry();
                CacheValue = TokenContext.Site.Id.ToString().Replace("-", "%2D");
            }
            return CacheValue;
        }
    }
}