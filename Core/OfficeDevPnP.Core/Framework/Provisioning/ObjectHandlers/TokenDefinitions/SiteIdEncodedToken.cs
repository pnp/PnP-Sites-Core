using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{siteidencoded}",
        Description = "Returns the id of the current site",
        Example = "{siteidencoded}",
        Returns = "9188a794%2Dcfcf%2D48b6%2D9ac5%2Ddf2048e8aa5d")]
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
                TokenContext.Load(TokenContext.Web, w => w.Id);
                TokenContext.ExecuteQueryRetry();
                CacheValue = TokenContext.Web.Id.ToString().Replace("-", "%2D");
            }
            return CacheValue;
        }
    }
}