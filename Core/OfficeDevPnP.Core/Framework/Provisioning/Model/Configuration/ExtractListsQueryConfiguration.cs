using Newtonsoft.Json;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration
{

    public partial class ExtractConfiguration
    {
        public class ExtractListsQueryConfiguration
        {
            [JsonProperty("camlQuery")]
            public string CamlQuery { get; set; }

            [JsonProperty("rowLimit")]
            public int RowLimit { get; set; }



            [JsonProperty("viewFields")]
            public List<string> ViewFields { get; set; }
        }
    }
}
