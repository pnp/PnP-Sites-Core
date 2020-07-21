using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Json header control data that will be included in each client side web part (de-)serialization (data-sp-controldata)
    /// </summary>
    public class HeaderControlData
    {
        /// <summary>
        /// Gets or sets JsonProperty "id"
        /// </summary>
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "instanceId"
        /// </summary>
        [JsonProperty(PropertyName = "instanceId")]
        public string InstanceId { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "title"
        /// </summary>
        [JsonProperty(PropertyName = "title")]
        public string Title { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "description"
        /// </summary>
        [JsonProperty(PropertyName = "description")]
        public string Description { get; set; }

        [JsonProperty(PropertyName = "serverProcessedContent")]
        public string ServerProcessedContent { get; internal set; }

        /// <summary>
        /// Gets or sets JsonProperty "dataVersion"
        /// </summary>
        [JsonProperty(PropertyName = "dataVersion")]
        public string DataVersion { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "properties"
        /// </summary>
        [JsonProperty(PropertyName = "properties")]
        public string Properties { get; set; }

    }
#endif
}
