using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class GroupIdToken : TokenDefinition
    {
        private readonly int _groupId = 0;
        public GroupIdToken(Web web, string name, int groupId)
            : base(web, $"{{groupid:{Regex.Escape(name)}}}")
        {
            _groupId = groupId;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _groupId.ToString();
            }
            return CacheValue;
        }
    }
}