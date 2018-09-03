using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Base class representing the json control data that will be included in each client side control (de-)serialization (data-sp-controldata attribute)
    /// </summary>
    public class ClientSideCanvasControlData
    {
        /// <summary>
        /// Gets or sets JsonProperty "controlType"
        /// </summary>
        [JsonProperty(PropertyName = "controlType")]
        public int ControlType { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "id"
        /// </summary>
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "position"
        /// </summary>
        [JsonProperty(PropertyName = "position")]
        public ClientSideCanvasControlPosition Position { get; set; }
    }
#endif
}
