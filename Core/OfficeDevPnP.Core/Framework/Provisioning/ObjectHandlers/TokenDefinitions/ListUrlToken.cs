using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class ListUrlToken : TokenDefinition
    {
        private string _listUrl = null;
        public ListUrlToken(Web web, string name, string url)
            : base(web, string.Format("{{listurl:{0}}}", Regex.Escape(name)))
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