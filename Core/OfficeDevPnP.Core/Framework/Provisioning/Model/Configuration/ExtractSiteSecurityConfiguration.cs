using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration
{
    public partial class ExtractConfiguration
    {
        public class ExtractSiteSecurityConfiguration
        {
            [JsonProperty("includeSiteGroups")]
            public bool IncludeSiteGroups { get; set; }
        }
    }
}
