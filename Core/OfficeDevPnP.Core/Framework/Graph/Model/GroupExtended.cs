using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    public class GroupExtended : Group
    {
        [JsonProperty("owners@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] OwnersODataBind { get; set; }
        [JsonProperty("members@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] MembersODataBind { get; set; }
    }
}
