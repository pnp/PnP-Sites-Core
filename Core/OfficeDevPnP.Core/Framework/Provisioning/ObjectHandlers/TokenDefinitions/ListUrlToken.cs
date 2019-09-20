using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{listurl:[name]}",
     Description = "Returns a site relative url of the list given its name",
     Example = "{listid:My List}",
     Returns = "Lists/MyList")]
    internal class ListUrlToken : TokenDefinition
    {
        private string _listUrl = null;
        public ListUrlToken(Web web, string name, string url)
            : base(web, $"{{listurl:{Regex.Escape(name)}}}")
        {
            _listUrl = url;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _listUrl;
            }
            return CacheValue;
        }
    }
}