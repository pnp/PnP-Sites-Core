using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class FileUniqueIdEncodedToken : TokenDefinition
    {
        private readonly string _value = null;
        public FileUniqueIdEncodedToken(Web web, string siteRelativePath, Guid uniqueId)
            : base(web, $"{{fileuniqueidencoded:{Regex.Escape(siteRelativePath)}}}")
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

