using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Themes.Generator
{
    public class ThemeRules : IThemeRules
    {
        private Dictionary<string, IThemeSlotRule> _rules = new Dictionary<string, IThemeSlotRule>();

        public IThemeSlotRule this[string key]
        {
            get => _rules[key];
            set => _rules[key] = value;
        }

        public IEnumerator<String> GetEnumerator()
        {
            return _rules.Select(s => s.Key).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
