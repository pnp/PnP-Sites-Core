using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{sitename}",
      Description = "Returns the title of the current site",
      Example = "{sitename}",
      Returns = "My Company Portal")]
    internal class SiteTitleToken : TokenDefinition
    {
        public SiteTitleToken(Web web) : base(web, "{sitetitle}", "~sitename", "{sitename}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Load(TokenContext.Web, w => w.Title);
                TokenContext.ExecuteQueryRetry();
                CacheValue = TokenContext.Web.Title;
            }
            return CacheValue;
        }
    }
}