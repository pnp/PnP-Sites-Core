using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Fields
{
    public class ApplyFieldsConfiguration
    {
        [JsonProperty("provisionFieldsToSubWebs")]
        public bool ProvisionFieldsToSubWebs { get; set; }
    }
}
