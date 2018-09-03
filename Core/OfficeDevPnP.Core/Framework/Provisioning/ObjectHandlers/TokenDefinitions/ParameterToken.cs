using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{parameter:[parametername]}",
        Description = "Returns the value of a parameter defined in the template",
        Example = "{parameter:MyParameter}",
        Returns = "the value of the parameter")]
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