using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration
{

    public partial class ExtractConfiguration
    {
        public class ExtractPagesConfiguration
        {
            [JsonProperty("excludeAuthorInformation")]
            public bool ExcludeAuthorInformation { get; set; }

            [JsonProperty("includeAllClientSidePages")]
            public bool IncludeAllClientSidePages { get; set; }
        }
    }
}
