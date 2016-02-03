using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class LocalizationToken : TokenDefinition
    {
        private List<ResourceEntry> _resourceEntries;
        public LocalizationToken(Web web, string key, List<ResourceEntry> resourceEntries)
            : base(web, string.Format("{{loc:{0}}}", key), string.Format("{{localize:{0}}}", key), string.Format("{{localization:{0}}}", key), string.Format("{{resource:{0}}}", key), string.Format("{{res:{0}}}", key))
        {
            _resourceEntries = resourceEntries;
        }

        public override string GetReplaceValue()
        {
            var entry = _resourceEntries.FirstOrDefault(r => r.LCID == this.Web.Language);
            if (entry != null)
            {
                return entry.Value;
            }
            else { return ""; }

        }

        public List<ResourceEntry> ResourceEntries
        {
            get { return _resourceEntries;  }
        }
    }
}