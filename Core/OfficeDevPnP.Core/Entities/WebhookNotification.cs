#if !ONPREMISES
using Newtonsoft.Json;
using System;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Web hook notification model: this message is received when SharePoint "fires" a web hook 
    /// </summary>
    public class WebhookNotification
    {
        /// <summary>
        /// Webhook subscription id
        /// </summary>
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        /// <summary>
        /// An opaque string passed back to the client on all notifications. You can use this for validating notifications, tagging different subscriptions, or other reasons.
        /// </summary>
        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        /// <summary>
        /// The expiration date for subscription. The expiration date should not be more than 6 months. By default, subscriptions are set to expire 6 months from when they are created.
        /// </summary>
        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        /// <summary>
        /// The resource endpoint URL you are creating the subscription for. For example a SharePoint List API URL
        /// </summary>
        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        /// <summary>
        /// SharePoint tenant id
        /// </summary>
        [JsonProperty(PropertyName = "tenantId")]
        public string TenantId { get; set; }

        /// <summary>
        /// SharePoint site URL
        /// </summary>
        [JsonProperty(PropertyName = "siteUrl")]
        public string SiteUrl { get; set; }

        /// <summary>
        /// SharePoint web id
        /// </summary>
        [JsonProperty(PropertyName = "webId")]
        public string WebId { get; set; }
    }
}
#endif