using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class StorageEntityValueToken : TokenDefinition
    {
        private string _key;
        private string _value;
        public StorageEntityValueToken(Web web, string key, string value)
            : base(web, $"{{storageentityvalue:{Regex.Escape(key)}}}")
        {
            _value = value;
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}