using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Lists.Lists
{
    public class ExtractListsListsConfiguration
    {
        [JsonProperty("title")]
        public string Title { get; set; }

        [JsonProperty("includeItems")]
        public bool IncludeItems { get; set; }

        [JsonProperty("keyColumn")]
        public string KeyColumn { get; set; }

        [JsonProperty("updateBehavior")]
        public UpdateBehavior UpdateBehavior { get; set; }

        [JsonProperty("skipEmptyFields")]
        public bool SkipEmptyFields { get; set; }

        [JsonProperty("query")]
        public ExtractListsQueryConfiguration Query { get; set; } = new ExtractListsQueryConfiguration();

        [JsonProperty("removeExistingContentTypes")]
        public bool RemoveExistingContentTypes { get; set; }

    }
}
