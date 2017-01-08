using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
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