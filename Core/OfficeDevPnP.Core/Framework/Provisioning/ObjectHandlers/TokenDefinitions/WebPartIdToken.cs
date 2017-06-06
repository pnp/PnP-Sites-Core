using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class WebPartIdToken : TokenDefinition
    {
        private string _webpartId = null;
        public WebPartIdToken(Web web, string name, Guid webpartid)
            : base(web, $"{{webpartid:{Regex.Escape(name)}}}")
        {
            _webpartId = webpartid.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _webpartId;
            }
            return CacheValue;
        }
    }
}