using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{o365groupid:[groupname]}",
     Description = "Returns the id of an Office 365 Group",
     Example = "{o365groupid:CompanyManagement}",
     Returns = "6")]
    internal class O365GroupIdToken : TokenDefinition
    {
        private readonly string _groupId = string.Empty;
        public O365GroupIdToken(Web web, string name, string groupId)
            : base(web, $"{{o365groupid:{Regex.Escape(name)}}}")
        {
            _groupId = groupId;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _groupId;
            }
            return CacheValue;
        }
    }
}