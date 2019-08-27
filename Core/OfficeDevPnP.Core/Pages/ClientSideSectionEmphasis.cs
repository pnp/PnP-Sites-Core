using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
    public class ClientSideSectionEmphasis
    {
        [JsonProperty(PropertyName = "zoneEmphasis", NullValueHandling = NullValueHandling.Ignore)]
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
