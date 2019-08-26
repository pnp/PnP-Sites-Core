using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
    public class ClientSideSectionEmphasis
    {
        [JsonIgnore]
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

        [JsonProperty(PropertyName = "zoneEmphasis", NullValueHandling = NullValueHandling.Ignore)]
        public string ZoneEmphasisString { get; set; }
    }
}
