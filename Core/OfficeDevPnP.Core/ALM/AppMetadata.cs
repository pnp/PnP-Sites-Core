using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.ALM
{
    public class AppMetadata
    {
        [JsonProperty()]
        public Guid Id { get; internal set; }

        [JsonProperty()]
        public Version AppCatalogVersion { get; internal set; }
        [JsonProperty()]
        public bool CanUpgrade { get; internal set; }
        [JsonProperty()]
        public bool Deployed { get; internal set; }
        [JsonProperty()]
        public Version InstalledVersion { get; internal set; }
        [JsonProperty()]
        public bool IsClientSideSolution { get; internal set; }
        [JsonProperty()]
        public string Title { get; internal set; }
    }


}
