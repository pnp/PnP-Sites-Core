using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class ContentTypeIdToken : TokenDefinition
    {
        private string _contentTypeId = null;
        public ContentTypeIdToken(Web web, string name, string contenttypeid)
            : base(web, string.Format("{{contenttypeid:{0}}}", name))
        {
            _contentTypeId = contenttypeid;
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