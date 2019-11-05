using Newtonsoft.Json;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Lists.Lists
{

    public class ExtractListsQueryConfiguration
    {
        [JsonProperty("camlQuery")]
        public string CamlQuery { get; set; }

        [JsonProperty("rowLimit")]
        public int RowLimit { get; set; }

        [JsonProperty("viewFields")]
        public List<string> ViewFields { get; set; } = new List<string>();

        [JsonProperty("includeAttachments")]
        public bool IncludeAttachments { get; set; }

        [JsonProperty("pageSize")]
        public int PageSize { get; set; }
    }
}
