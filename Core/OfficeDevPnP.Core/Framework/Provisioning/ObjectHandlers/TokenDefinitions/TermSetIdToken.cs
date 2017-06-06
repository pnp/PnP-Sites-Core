using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class TermSetIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public TermSetIdToken(Web web, string groupName, string termsetName, Guid id)
            : base(web, $"{{termsetid:{Regex.Escape(groupName)}:{Regex.Escape(termsetName)}}}")
        {
            _value = id.ToString();
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