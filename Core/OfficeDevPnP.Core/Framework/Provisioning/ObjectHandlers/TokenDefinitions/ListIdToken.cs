using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
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