using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class WebPartIdToken : TokenDefinition
    {
        private string _webpartId = null;
        public WebPartIdToken(Web web, string name, Guid webpartid)
            : base(web, string.Format("{{webpartid:{0}}}", name))
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