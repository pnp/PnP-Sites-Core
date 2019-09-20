using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{pageuniqueidencoded:[siterelativepath]}",
     Description = "Returns the HTML safe encoded id of a client side page that is being provisioned through the current template",
     Example = "{pageuniqueidencoded:SitePages/Home.aspx}",
     Returns = "767bc144%2De605%2D4d8c%2D885a%2D3a980feb39c6")]
    internal class PageUniqueIdEncodedToken : TokenDefinition
    {
        private readonly string _value = null;
        public PageUniqueIdEncodedToken(Web web, string siteRelativePath, Guid uniqueId)
            : base(web, $"{{pageuniqueidencoded:{Regex.Escape(siteRelativePath)}}}")
        {
            _value = uniqueId.ToString().Replace("-", "%2D");
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}

