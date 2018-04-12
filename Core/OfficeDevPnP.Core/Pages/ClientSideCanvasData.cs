using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
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
    }
#endif
}
