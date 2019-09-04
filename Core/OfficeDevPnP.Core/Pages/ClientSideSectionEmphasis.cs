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
            get
            {
                if (!string.IsNullOrWhiteSpace(ZoneEmphasisString) && int.TryParse(ZoneEmphasisString, out int result))
                {
                    return result;
                }
                return 0;
            }
            set { ZoneEmphasisString = value.ToString(); }
        }

        [JsonIgnore]
        public string ZoneEmphasisString { get; set; }
    }
}
