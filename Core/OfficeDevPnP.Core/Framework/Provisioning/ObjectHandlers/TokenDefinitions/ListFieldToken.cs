using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class ListFieldToken : TokenDefinition
    {
        private string _fieldId = null;
        public ListFieldToken(Web web, string listName, string fieldName, Guid listid)
            : base(web, string.Format("{{listFieldName:{0}:{1}}}", Regex.Escape(listName), Regex.Escape(fieldName)))
        {
            _fieldId = listid.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _fieldId;
            }
            return CacheValue;
        }
    }
}