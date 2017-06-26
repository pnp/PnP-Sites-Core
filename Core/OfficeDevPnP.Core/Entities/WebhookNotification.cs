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
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        [JsonProperty(PropertyName = "tenantId")]
        public string TenantId { get; set; }

        [JsonProperty(PropertyName = "siteUrl")]
        public string SiteUrl { get; set; }

        [JsonProperty(PropertyName = "webId")]
        public string WebId { get; set; }
    }
}
#endif