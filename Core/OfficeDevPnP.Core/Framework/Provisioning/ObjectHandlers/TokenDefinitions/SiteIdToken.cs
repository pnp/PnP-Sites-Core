using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{siteid}",
       Description = "Returns the id of the current site",
       Example = "{siteid}",
       Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
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
                TokenContext.Load(TokenContext.Web, w => w.Id);
                TokenContext.ExecuteQueryRetry();
                CacheValue = TokenContext.Web.Id.ToString();
            }
            return CacheValue;
        }
    }
}