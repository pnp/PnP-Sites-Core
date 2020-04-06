using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    public class GroupCreationRequest
    {
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("groupTypes")]
        public string[] GroupTypes { get; set; } = new string[] { "Unified" };

        [JsonProperty("mailEnabled")]
        public bool MailEnabled { get; set; } = true;

        [JsonProperty("securityEnabled")]
        public bool SecurityEnabled { get; set; } = false;

        [JsonProperty("mailNickname")]
        public string MailNickname { get; set; }

        [JsonProperty("owners@odata.bind")]
        public string[] Owners { get; set; }

        [JsonProperty("members@odata.bind")]
        public string[] Members { get; set; }
    }
}
