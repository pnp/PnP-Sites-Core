using Newtonsoft.Json;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Lists.Lists
{

    public class ExtractQueryConfiguration
    {
        [JsonProperty("camlQuery")]
        public string CamlQuery { get; set; }

        [JsonProperty("rowLimit")]
        public int RowLimit { get; set; }

        [JsonProperty("viewFields")]
        public List<string> ViewFields { get; set; }

        [JsonProperty("includeAttachments")]
        public bool IncludeAttachments { get; set; }

        [JsonProperty("pageSize")]
        public int PageSize { get; set; }
    }
}
