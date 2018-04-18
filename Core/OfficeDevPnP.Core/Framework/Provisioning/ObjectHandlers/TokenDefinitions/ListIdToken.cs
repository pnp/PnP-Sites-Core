using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{listid:[name]}",
     Description = "Returns a id of the list given its name",
     Example = "{listid:My List}",
     Returns = "f2cd6d5b-1391-480e-a3dc-7f7f96137382")]
    internal class ListIdToken : TokenDefinition
    {
        private string _listId = null;

        public ListIdToken(Web web, string name, Guid listid)
            : base(web, $"{{listid:{Regex.Escape(name)}}}")
        {
            _listId = listid.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _listId;
            }
            return CacheValue;
        }
    }
}