using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Utilities.Themes.Palettes;

namespace OfficeDevPnP.Core.Utilities.Themes.Generator
{
    public class ThemeSlotRule : IThemeSlotRule
    {
        public string name { get; set; }
        public IColor color { get; set; }
        public string value { get; set; }
        public IThemeSlotRule inherits { get; set; }
        public Shade asShade { get; set; }
        public bool isBackgroundShade { get; set; }
        public bool isCustomized { get; set; }
        public List<IThemeSlotRule> dependentRules { get; set; }
    }
}
