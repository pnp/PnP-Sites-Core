using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Class representing the json control data that will describe a control versus the zones and sections on a page
    /// </summary>
    public class ClientSideCanvasControlPosition : ClientSideCanvasPosition
    {
        /// <summary>
        /// Gets or sets JsonProperty "controlIndex"
        /// </summary>
        [JsonProperty(PropertyName = "controlIndex")]
        public float ControlIndex { get; set; }
    }
#endif
}
