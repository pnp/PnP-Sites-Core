using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    public static class StringUtilities
    {
        public static string[] Split(this string input, string separator)
        {
            var splitRegex = new Regex(
                Regex.Escape(separator),
                RegexOptions.Singleline | RegexOptions.Compiled
                );

            return splitRegex.Split(input);
        }
    }
}
