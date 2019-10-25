using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Tenant.Sequence
{
    public class ExtractSequenceConfiguration
    {
        [JsonProperty("siteUrls")]
        public List<string> SiteUrls { get; set; } = new List<string>();

        [JsonProperty("maxSubsiteDepth")]
        public int MaxSubsiteDepth { get; set; }

        [JsonProperty("includeJoinedSites")]
        public bool IncludeJoinedSites { get; set; }

        [JsonProperty("includeSubsites")]
        public bool IncludeSubsites { get; set; }
    }
}
