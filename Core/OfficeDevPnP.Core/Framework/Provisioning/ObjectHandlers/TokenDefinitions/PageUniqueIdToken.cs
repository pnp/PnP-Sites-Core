using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class PageUniqueIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public PageUniqueIdToken(Web web, string siteRelativePath, Guid uniqueId)
            : base(web, $"{{pageuniqueid:{Regex.Escape(siteRelativePath)}}}")
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

