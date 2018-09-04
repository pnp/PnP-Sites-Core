using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{webname}",
      Description = "Returns the name part of the URL of the Server Relative URL of the Web",
      Example = "{webname}",
      Returns = "MyWeb")]
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