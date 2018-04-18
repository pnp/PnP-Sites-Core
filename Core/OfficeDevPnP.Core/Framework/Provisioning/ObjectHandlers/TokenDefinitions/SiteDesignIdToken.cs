using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitedesignid:[designtitle]}",
        Description = "Returns the id of the given site design",
        Example = "{sitedesignid:My Site Design}",
        Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
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