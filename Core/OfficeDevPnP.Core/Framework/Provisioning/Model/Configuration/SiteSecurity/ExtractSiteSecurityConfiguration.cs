using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.SiteSecurity
{
    public class ExtractConfiguration
    {
        [JsonProperty("includeSiteGroups")]
        public bool IncludeSiteGroups { get; set; }
    }
}
