using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{fileuniqueidencoded:[siteRelativePath]}",
     Description = "Returns the html safe encoded unique id of a file which is being provisioned by the current template.",
     Example = "{fileuniqueid:/sitepages/home.aspx}",
     Returns = "f2cd6d5b%2D1391%2D480e%2Da3dc%2D7f7f96137382")]
    internal class FileUniqueIdEncodedToken : TokenDefinition
    {
        private readonly string _value = null;
        public FileUniqueIdEncodedToken(Web web, string siteRelativePath, Guid uniqueId)
            : base(web, $"{{fileuniqueidencoded:{Regex.Escape(siteRelativePath)}}}")
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

