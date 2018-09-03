using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{webpartid:[webpartname]}",
      Description = "Returns the id of a webpart that is being provisioned to a page through a template",
      Example = "{webpartid:mywebpart}",
      Returns = "66e2b037-f749-402d-90b2-afd643850c26")]
    internal class WebPartIdToken : TokenDefinition
    {
        private string _webpartId = null;
        public WebPartIdToken(Web web, string name, Guid webpartid)
            : base(web, $"{{webpartid:{Regex.Escape(name)}}}")
        {
            _webpartId = webpartid.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _webpartId;
            }
            return CacheValue;
        }
    }
}