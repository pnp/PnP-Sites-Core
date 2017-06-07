using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Represents the payload of a Http message
    /// </summary>
    public class WebhookSubscription
    {
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "notificationUrl")]
        public string NotificationUrl { get; set; }

        [JsonProperty(PropertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
        public string Resource { get; set; }
        
    }
}
