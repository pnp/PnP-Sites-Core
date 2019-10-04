using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Base class representing the json control data that will be included in each client side control (de-)serialization (data-sp-controldata attribute)
    /// </summary>
    public class ClientSideCanvasData
    {
        /// <summary>
        /// Gets or sets JsonProperty "position"
        /// </summary>
        [JsonProperty(PropertyName = "position")]
        public ClientSideCanvasPosition Position { get; set; }

        [JsonProperty(PropertyName = "emphasis", NullValueHandling = NullValueHandling.Ignore)]
        public ClientSideSectionEmphasis Emphasis { get; set; }

        [JsonProperty(PropertyName = "pageSettingsSlice", NullValueHandling = NullValueHandling.Ignore)]
        public ClientSidePageSettingsSlice PageSettingsSlice {get ;set;}
    }
#endif
}
