using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{viewid:[listname],[viewname]}",
     Description = "Returns a id of the view given its name for a given list",
     Example = "{viewid:My List,My View}",
     Returns = "f2cd6d5b-1391-480e-a3dc-7f7f96137382")]
    internal class ListViewIdToken : TokenDefinition
    {
        private string _viewId = null;

        public ListViewIdToken(Web web, string listTitle, string viewTitle, Guid viewId)
            : base(web, $"{{viewid:{Regex.Escape(listTitle)},{Regex.Escape(viewTitle)}}}")
        {
            _viewId = viewId.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _viewId;
            }
            return CacheValue;
        }
    }
}