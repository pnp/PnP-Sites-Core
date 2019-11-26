using Newtonsoft.Json;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.ContentTypes
{
    public class ExtractContentTypeConfiguration
    {
        [JsonProperty("Groups")]
        public List<string> Groups { get; set; } = new List<string>();

        [JsonProperty("IncludeFromSyndication")]
        public bool ExcludeFromSyndication { get; set; }
    }
}
