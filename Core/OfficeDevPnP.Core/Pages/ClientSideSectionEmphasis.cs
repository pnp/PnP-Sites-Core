using Newtonsoft.Json;
using OfficeDevPnP.Core.Utilities.JsonConverters;

namespace OfficeDevPnP.Core.Pages
{
    public class ClientSideSectionEmphasis
    {
        [JsonProperty(PropertyName = "zoneEmphasis", NullValueHandling = NullValueHandling.Ignore)]
        [JsonConverter(typeof(EmphasisJsonConverter))]
        public int ZoneEmphasis
        {
            get; set;
        }
    }
}
