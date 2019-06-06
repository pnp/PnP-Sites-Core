using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Pages
{
    public class ClientSideSectionEmphasis
    {
        [JsonProperty(PropertyName = "zoneEmphasis")]
        public int ZoneEmphasis { get; set; }
    }
}
