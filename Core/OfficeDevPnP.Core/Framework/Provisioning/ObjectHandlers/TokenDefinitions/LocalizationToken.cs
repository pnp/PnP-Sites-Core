using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class LocalizationToken : TokenDefinition
    {
        private List<ResourceEntry> _resourceEntries;
        public LocalizationToken(Web web, string key, List<ResourceEntry> resourceEntries)
            : base(web, $"{{loc:{Regex.Escape(key)}}}", $"{{localize:{Regex.Escape(key)}}}", $"{{localization:{Regex.Escape(key)}}}", $"{{resource:{Regex.Escape(key)}}}", $"{{res:{Regex.Escape(key)}}}")
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