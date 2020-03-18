using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{listContentTypeId:[listname],[contentTypeName]}",
     Description = "Returns an id of the content type given its name for a given list",
     Example = "{listContentTypeId:My List,Folder}",
     Returns = "0x010100F0D7B2FF0128AD459168DFA77A2A1BD0")]
    [TokenDefinitionDescription(
     Token = "{listContentTypeId:[listname],[contentTypeId]}",
     Description = "Returns an id of the content type given its direct parent id for a given list",
     Example = "{listContentTypeId:My List,Folder}",
     Returns = "0x0101")]
    internal class ListContentTypeIdToken : TokenDefinition
    {
        private string _contentTypeId = null;

        public ListContentTypeIdToken(Web web, string listTitle, string contentTypeName, ContentTypeId contentTypeId)
            : base(web, $"{{listcontenttypeid:{Regex.Escape(listTitle)},{Regex.Escape(contentTypeName)}}}")
        {
            _contentTypeId = contentTypeId.StringValue;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _contentTypeId;
            }
            return CacheValue;
        }
    }
}