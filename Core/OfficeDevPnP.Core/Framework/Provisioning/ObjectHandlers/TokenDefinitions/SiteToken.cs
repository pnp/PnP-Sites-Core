using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{site}",
      Description = "Returns the server relative url of the current site",
      Example = "{site}",
      Returns = "/sites/mysitecollection/mysite")]
    internal class SiteToken : TokenDefinition
    {
        public SiteToken(Web web)
            : base(web, "~site", "{site}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Load(TokenContext.Web, w => w.ServerRelativeUrl);
                TokenContext.ExecuteQueryRetry();
                CacheValue = TokenContext.Web.ServerRelativeUrl.TrimEnd('/');
            }
            return CacheValue;
        }
    }
}