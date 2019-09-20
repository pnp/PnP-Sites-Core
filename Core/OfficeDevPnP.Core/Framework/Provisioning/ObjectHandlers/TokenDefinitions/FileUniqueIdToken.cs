using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{fileuniqueid:[siteRelativePath]}",
     Description = "Returns the unique id of a file which is being provisioned by the current template.",
     Example = "{fileuniqueid:/sitepages/home.aspx}",
     Returns = "f2cd6d5b-1391-480e-a3dc-7f7f96137382")]
    internal class FileUniqueIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public FileUniqueIdToken(Web web, string siteRelativePath, Guid uniqueId)
            : base(web, $"{{fileuniqueid:{Regex.Escape(siteRelativePath)}}}")
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

