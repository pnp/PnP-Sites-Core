using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    public class ClientSidePageSettingsSlice
    {
        [JsonProperty(PropertyName = "isDefaultDescription", NullValueHandling = NullValueHandling.Ignore)]
        public bool? IsDefaultDescription { get; set; }

        [JsonProperty(PropertyName = "isDefaultThumbnail", NullValueHandling = NullValueHandling.Ignore)]
        public bool? IsDefaultThumbnail { get; set; }
    }
#endif
}
