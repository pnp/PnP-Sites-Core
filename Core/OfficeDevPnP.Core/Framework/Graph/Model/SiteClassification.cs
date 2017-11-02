using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    public class SiteClassificationSettings
    {
        public string UsageGuidelinesUrl { get; set; } = "";
        public List<string> Classifications { get; set; }
        public string DefaultClassification { get; set; } = "";
    }
}
