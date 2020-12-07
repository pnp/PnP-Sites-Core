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
     Example = "{listContentTypeId:My List,Document}",
     Returns = "0x010100F0D7B2FF0128AD459168DFA77A2A1BD0")]
    [TokenDefinitionDescription(
     Token = "{listContentTypeId:[listname],[contentTypeId]}",
     Description = "Returns an id of the content type given its direct parent id for a given list",
     Example = "{listContentTypeId:My List,0x0101}",
     Returns = "0x010100F0D7B2FF0128AD459168DFA77A2A1BD0")]
    internal class ListContentTypeIdToken : TokenDefinition
    {
        private string _contentTypeId = null;
        private const string tokenPrefix = "listcontenttypeid";

        [Obsolete("Use ListContentTypeIdToken(Web, string, ContentType) instead")]
        public ListContentTypeIdToken(Web web, string listTitle, string contentTypeName, ContentTypeId contentTypeId)
            : base(web, $"{{listcontenttypeid:{Regex.Escape(listTitle)},{Regex.Escape(contentTypeName)}}}")
        {
            _contentTypeId = contentTypeId.StringValue;
        }

        public ListContentTypeIdToken(Web web, string listTitle, ContentType contentType)
            : base(web, 
                  CreateToken(listTitle, contentType.Id),
                  CreateToken(listTitle, contentType.Name))
        {
            _contentTypeId = contentType.Id.StringValue;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _contentTypeId;
            }
            return CacheValue;
        }

        /// <summary>
        /// Creates a token for the specified list title and list content type name
        /// </summary>
        /// <param name="listTitle">Title for the list</param>
        /// <param name="contentTypeName">Name of the list content type</param>
        /// <returns>List content type token. Example: "{listContentTypeId:My List,Document}"</returns>
        public static string CreateToken(string listTitle, string contentTypeName)
        {
            return $"{{{tokenPrefix}:{Regex.Escape(listTitle)},{Regex.Escape(contentTypeName)}}}";
        }

        /// <summary>
        /// Creates a token for the specified list title and list content type id
        /// </summary>
        /// <param name="listTitle">Title for the list</param>
        /// <param name="id">Content type id of the list content type</param>
        /// <returns>List content type token. Example: "{listContentTypeId:My List,0x0101}"</returns>
        public static string CreateToken(string listTitle, ContentTypeId id)
        {
            return $"{{{tokenPrefix}:{Regex.Escape(listTitle)},{Regex.Escape(id.GetParentIdValue())}}}";
        }
    }
}