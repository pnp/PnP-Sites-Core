using System;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Defines a Microsoft Graph Subscription
    /// </summary>
    public class Subscription
    {
        /// <summary>
        /// Indicates the type(s) of change(s) in the subscribed resource that will raise a notification
        /// </summary>
        public Enums.GraphSubscriptionChangeType? ChangeType { get; set; }

        /// <summary>
        /// The URL of the endpoint that will receive the notifications. This URL must make use of the HTTPS protocol.
        /// </summary>
        public string NotificationUrl { get; set; }

        /// <summary>
        /// Specifies the resource that will be monitored for changes. Do not include the base URL (https://graph.microsoft.com/v1.0/). See the possible resource path values for each supported resource.
        /// </summary>
        public string Resource { get; set; }

        /// <summary>
        /// Identifier of the application used to create the subscription
        /// </summary>
        public string ApplicationId { get; set; }

        /// <summary>
        /// Specifies the date and time when the webhook subscription expires. The time is in UTC, and can be an amount of time from subscription creation that varies for the resource subscribed to. See https://docs.microsoft.com/graph/api/resources/subscription#maximum-length-of-subscription-per-resource-type for maximum supported subscription length of time.
        /// </summary>
        public DateTimeOffset? ExpirationDateTime { get; set; }

        /// <summary>
        /// 	Unique identifier for the subscription
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Specifies the value of the clientState property sent by the service in each notification. The maximum length is 128 characters. The client can check that the notification came from the service by comparing the value of the clientState property sent with the subscription with the value of the clientState property received with each notification.
        /// </summary>
        public string ClientState { get; set; }

        /// <summary>
        /// Identifier of the user or service principal that created the subscription. If the app used delegated permissions to create the subscription, this field contains the id of the signed-in user the app called on behalf of. If the app used application permissions, this field contains the id of the service principal corresponding to the app
        /// </summary>
        public string CreatorId { get; set; }

        /// <summary>
        /// Specifies the latest version of Transport Layer Security (TLS) that the notification endpoint, specified by notificationUrl, supports. The possible values are: v1_0, v1_1, v1_2, v1_3. For subscribers whose notification endpoint supports a version lower than the currently recommended version(TLS 1.2), specifying this property by a set timeline allows them to temporarily use their deprecated version of TLS before completing their upgrade to TLS 1.2. For these subscribers, not setting this property per the timeline would result in subscription operations failing.
        /// </summary>
        public string LatestSupportedTlsVersion { get; set; }
    }
}
