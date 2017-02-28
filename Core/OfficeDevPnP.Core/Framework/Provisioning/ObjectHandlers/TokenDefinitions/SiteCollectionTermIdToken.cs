using System;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    public class SiteCollectionTermIdToken : TokenDefinition
    {
        private readonly string _termId;

        public SiteCollectionTermIdToken(Web ctxWeb, string termsetName, string termPath, Guid termId) : base(ctxWeb, $"{{sitecollectiontermid:{Regex.Escape(termsetName)}:{Regex.Escape(termPath)}}}")
        {
            _termId = termId.ToString("D");
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _termId;
            }
            return CacheValue;
        }
    }
}
