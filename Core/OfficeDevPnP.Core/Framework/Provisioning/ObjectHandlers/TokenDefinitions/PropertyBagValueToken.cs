using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{propertybagvalue:[key]}",
        Description = "Returns the value of a propertybag value",
        Example = "{propertybagvalue:MyKey}",
        Returns = "the value of the propertybag value defined by the key")]
    internal class PropertyBagValueToken : TokenDefinition
    {
        private readonly string _value = null;
        public PropertyBagValueToken(Web web, string name, string value)
            : base(web, $"{{propertybagvalue:{Regex.Escape(name)}}}")
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