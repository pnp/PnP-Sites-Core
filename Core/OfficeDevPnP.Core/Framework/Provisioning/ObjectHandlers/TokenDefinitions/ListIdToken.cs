using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{listid:[name]}",
     Description = "Returns a id of the list given its name",
     Example = "{listid:My List}",
     Returns = "f2cd6d5b-1391-480e-a3dc-7f7f96137382")]
    internal class ListIdToken : VolatileTokenDefinition
    {
        private string _listId = null;
        private string _name = null;
        public ListIdToken(Web web, string name, Guid listid)
            : base(web, $"{{listid:{Regex.Escape(name)}}}")
        {
            if (listid == Guid.Empty)
            {
                // on demand loading
                _name = name;
            }
            else
            {
                _listId = listid.ToString();
            }
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                if (_listId != null)
                {
                    CacheValue = _listId;
                }
                else
                {
                    var list = TokenContext.Web.Lists.GetByTitle(_name);
                    TokenContext.Load(list, l => l.Id);
                    TokenContext.ExecuteQueryRetry();
                    _listId = list.Id.ToString();
                    CacheValue = list.Id.ToString();
                }
            }
            return CacheValue;
        }
    }
}