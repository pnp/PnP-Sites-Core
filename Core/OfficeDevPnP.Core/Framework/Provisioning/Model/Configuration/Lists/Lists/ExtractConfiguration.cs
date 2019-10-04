using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Lists.Lists
{
    public class ExtractConfiguration
    {
        [JsonProperty("title")]
        public string Title { get; set; }

        [JsonProperty("includeItems")]
        public bool IncludeItems { get; set; }

        [JsonProperty("skipEmptyFields")]
        public bool SkipEmptyFields { get; set; }

        [JsonProperty("query")]
        public ExtractQueryConfiguration Query { get; set; }

        [JsonProperty("removeExistingContentTypes")]
        public bool RemoveExistingContentTypes { get; set; }

    }
}
