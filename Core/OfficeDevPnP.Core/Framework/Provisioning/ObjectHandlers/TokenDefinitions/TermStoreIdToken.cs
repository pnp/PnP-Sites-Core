using Microsoft.SharePoint.Client;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class TermStoreIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public TermStoreIdToken(Web web, string storeName, Guid id)
            : base(web, string.Format("{{termstoreid:{0}}}", storeName))
        {
            _value = id.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}