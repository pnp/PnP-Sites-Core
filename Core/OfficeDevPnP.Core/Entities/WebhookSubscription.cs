#if !ONPREMISES
using Newtonsoft.Json;
using System;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Represents the payload of a Http message
    /// </summary>
    public class WebhookSubscription
    {
        /// <summary>
        /// Webhook subscription id
        /// </summary>
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }
        /// <summary>
        /// An opaque string passed back to the client on all notifications. You can use this for validating notifications, tagging different subscriptions, or other reasons.
        /// </summary>
        [JsonProperty(PropertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
        public string ClientState { get; set; }
        /// <summary>
        /// The expiration date for subscription. The expiration date should not be more than 6 months. By default, subscriptions are set to expire 6 months from when they are created.
        /// </summary>
        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }
        /// <summary>
        /// Webhook notification URL
        /// </summary>
        [JsonProperty(PropertyName = "notificationUrl", NullValueHandling = NullValueHandling.Ignore)]
        public string NotificationUrl { get; set; }
        /// <summary>
        /// The resource endpoint URL you are creating the subscription for. For example a SharePoint List API URL
        /// </summary>
        [JsonProperty(PropertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
        public string Resource { get; set; }
        
    }
}
#endif
