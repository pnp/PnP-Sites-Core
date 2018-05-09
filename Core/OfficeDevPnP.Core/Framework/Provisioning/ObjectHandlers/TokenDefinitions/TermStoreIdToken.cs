using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{termstoreid:[storename]}",
      Description = "Returns the id of a term store given its name",
      Example = "{termstoreid:MyTermStore}",
      Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
    internal class TermStoreIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public TermStoreIdToken(Web web, string storeName, Guid id)
            : base(web, $"{{termstoreid:{Regex.Escape(storeName)}}}")
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