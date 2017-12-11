using Microsoft.SharePoint.Client;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class AppPackageIdToken : TokenDefinition
    {
        private string _appPackageId = null;

        public AppPackageIdToken(Web web, string name, Guid appPackageId)
            : base(web, $"{{apppackageid:{Regex.Escape(name)}}}")
        {
            _appPackageId = appPackageId.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _appPackageId;
            }
            return CacheValue;
        }
    }
}