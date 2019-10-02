using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration
{

    public partial class ExtractConfiguration
    {
        public class ExtractListsListsConfiguration
        {
            [JsonProperty("title")]
            public string Title { get; set; }

            [JsonProperty("includeItems")]
            public bool IncludeItems { get; set; }

            [JsonProperty("skipEmptyFields")]
            public bool SkipEmptyFields { get; set; }

            [JsonProperty("query")]
            public ExtractListsQueryConfiguration Query { get; set; }

        }
    }
}
