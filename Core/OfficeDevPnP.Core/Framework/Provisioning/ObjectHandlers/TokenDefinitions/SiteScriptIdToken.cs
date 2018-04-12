using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class SiteScriptIdToken : TokenDefinition
    {
        private Guid _scriptId;
        public SiteScriptIdToken(Web web, string scriptTitle, Guid scriptId)
            : base(web, $"{{sitescriptid:{Regex.Escape(scriptTitle)}}}")
        {
            _scriptId = scriptId;
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                CacheValue = _scriptId.ToString();
            }
            return CacheValue;
        }
    }
}