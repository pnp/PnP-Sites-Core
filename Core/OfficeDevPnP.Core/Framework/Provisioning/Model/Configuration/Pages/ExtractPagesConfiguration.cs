using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Pages
{
    public class ExtractConfiguration
    {
        [JsonProperty("excludeAuthorInformation")]
        public bool ExcludeAuthorInformation { get; set; }

        [JsonProperty("includeAllClientSidePages")]
        public bool IncludeAllClientSidePages { get; set; }
    }
}
