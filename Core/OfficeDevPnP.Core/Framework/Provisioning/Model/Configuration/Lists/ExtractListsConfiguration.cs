using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Lists.Lists;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Lists
{
    public class ExtractListsConfiguration
    {
        [JsonProperty("includeHiddenLists")]
        public bool IncludeHiddenLists { get; set; }

        [JsonProperty("lists")]
        public List<Lists.ExtractListsListsConfiguration> Lists { get; set; } = new List<Lists.ExtractListsListsConfiguration>();

        public bool HasLists
        {
            get
            {
                return this.Lists != null && this.Lists.Count > 0;
            }
        }
    }
}
