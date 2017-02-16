using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class ParameterToken : TokenDefinition
    {
        private readonly string _value = null;
        public ParameterToken(Web web, string name, string value)
            : base(web, $"{{parameter:{Regex.Escape(name)}}}", $"{{\\${Regex.Escape(name)}}}")
        {
            _value = value;
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