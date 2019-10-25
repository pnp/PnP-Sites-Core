using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Tenant
{
    public class ExtractTenantConfiguration
    {
        /// <summary>
        /// If defined will extract site collections as defined in the SiteUrls array
        /// </summary>
        [JsonProperty("sequence")]
        public Sequence.ExtractSequenceConfiguration Sequence { get; set; }

        /// <summary>
        /// If defined will extract teams as defined
        /// </summary>
        [JsonProperty("teams")]
        public Teams.ExtractTeamsConfiguration Teams { get; set; }
    }
}
