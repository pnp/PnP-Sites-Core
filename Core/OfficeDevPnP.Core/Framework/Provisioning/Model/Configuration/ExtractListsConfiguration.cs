using Newtonsoft.Json;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration
{

    public partial class ExtractConfiguration
    {
        public class ExtractListsConfiguration
        {
            [JsonProperty("includeHiddenLists")]
            public bool IncludeHiddenLists { get; set; }

            [JsonProperty("lists")]
            public List<ExtractListsListsConfiguration> Lists { get; set; }

        }
    }
}
