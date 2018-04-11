using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteDesignIdToken : TokenDefinition
    {
        private Guid _designId;
        public SiteDesignIdToken(Web web, string designTitle, Guid designId)
            : base(web, $"{{sitedesignid:{Regex.Escape(designTitle)}}}")
        {
            _designId = designId;
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                CacheValue = _designId.ToString();
            }
            return CacheValue;
        }
    }
}