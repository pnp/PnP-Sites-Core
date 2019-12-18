using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Taxonomy
{
    public class ExtractTaxonomyConfiguration
    {
        [JsonProperty("includeSecurity")]
        public bool IncludeSecurity { get; set; }

        [JsonProperty("includeSiteCollectionTermGroup")]
        public bool IncludeSiteCollectionTermGroup { get; set; }

        [JsonProperty("includeAllTermGroups")]
        public bool IncludeAllTermGroups { get; set; }
    }
}
