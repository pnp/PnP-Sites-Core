using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Represents settings with regards to Site Classifications
    /// </summary>
    public class SiteClassificationsSettings
    {
        /// <summary>
        /// The URL pointing to usage guidelines with regards to site classifications
        /// </summary>
        public string UsageGuidelinesUrl { get; set; } = "";

        /// <summary>
        /// A list of classifications that was retrieved or should be applied.
        /// </summary>
        public List<string> Classifications { get; set; }

        /// <summary>
        /// The default classification to use. Notice that when applying or updating the value specified should be present in the Classifications list of values.
        /// </summary>
        public string DefaultClassification { get; set; } = "";
    }
}
