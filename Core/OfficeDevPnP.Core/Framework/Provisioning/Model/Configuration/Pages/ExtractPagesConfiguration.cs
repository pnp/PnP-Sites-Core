using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Pages
{
    public class ExtractPagesConfiguration
    {
        [JsonProperty("excludeAuthorInformation")]
        public bool ExcludeAuthorInformation { get; set; }

        [JsonProperty("includeAllClientSidePages")]
        public bool IncludeAllClientSidePages { get; set; }
    }
}
