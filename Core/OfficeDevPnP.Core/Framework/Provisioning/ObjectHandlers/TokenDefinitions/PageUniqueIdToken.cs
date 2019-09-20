using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
   Token = "{pageuniqueid:[siterelativepath]}",
   Description = "Returns the id of a client side page that is being provisioned through the current template",
   Example = "{pageuniqueid:SitePages/Home.aspx}",
   Returns = "767bc144-e605-4d8c-885a-3a980feb39c6")]
    internal class PageUniqueIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public PageUniqueIdToken(Web web, string siteRelativePath, Guid uniqueId)
            : base(web, $"{{pageuniqueid:{Regex.Escape(siteRelativePath)}}}")
        {
            _value = uniqueId.ToString();
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

