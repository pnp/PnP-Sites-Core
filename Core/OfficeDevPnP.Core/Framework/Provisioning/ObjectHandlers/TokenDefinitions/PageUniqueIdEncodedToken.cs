using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class PageUniqueIdEncodedToken : TokenDefinition
    {
        private readonly string _value = null;
        public PageUniqueIdEncodedToken(Web web, string siteRelativePath, Guid uniqueId)
            : base(web, $"{{pageuniqueidencoded:{Regex.Escape(siteRelativePath)}}}")
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

