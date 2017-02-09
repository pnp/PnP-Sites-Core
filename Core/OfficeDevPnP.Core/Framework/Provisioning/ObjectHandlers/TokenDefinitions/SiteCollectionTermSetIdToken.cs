using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteCollectionTermSetIdToken : TokenDefinition
    {
        private string _value;

        public SiteCollectionTermSetIdToken(Web web, string termsetName, Guid id)
            : base(web, $"{{sitecollectiontermsetid:{Regex.Escape(termsetName)}}}")
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