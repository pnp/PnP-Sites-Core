#if !ONPREMISES
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities.Async;
using OfficeDevPnP.Core.Utilities.Webhooks;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// The list of Hookable Resource Types
    /// </summary>
    public enum WebHookResourceType
    {
        /// <summary>
        /// List web hooks
        /// </summary>
        List,
    }

    /// <summary>
    /// Class containing utility methods to manage Webhook on a SharePoint resource
    /// Adapted from https://github.com/SharePoint/sp-dev-samples/blob/master/Samples/WebHooks.List/SharePoint.WebHooks.Common/WebHookManager.cs
    /// </summary>
    internal static class WebhookUtility
    {
        private const string SubscriptionsUrlPart = "subscriptions";
        private const string ListIdentifierFormat = @"{0}/_api/web/lists(guid'{1}')";
        public const int MaximumValidityInMonths = 6;
        public const int ExpirationDateTimeMaxDays = 180;

        /// <summary>
        /// The effective maximum expiration datetime
        /// </summary>
        public static DateTime MaxExpirationDateTime
        {
            get
            {
                return DateTime.UtcNow.AddDays(ExpirationDateTimeMaxDays);
            }
        }

        /// <summary>
        /// Add a Webhook subscription to a SharePoint resource
        /// </summary>
        /// <param name="webUrl">Url of the site holding the list</param>
        /// <param name="resourceType">The type of Hookable SharePoint resource</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <param name="context">ClientContext instance to use for authentication</param>
        /// <param name="subscription">The Webhook subscription to add</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when expiration date is out of valid range.</exception>
        /// <returns>The added subscription object</returns>
        public static async Task<WebhookSubscription> AddWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string accessToken, ClientContext context, WebhookSubscription subscription)
        {
            if (!ValidateExpirationDateTime(subscription.ExpirationDateTime))
                throw new ArgumentOutOfRangeException(nameof(subscription.ExpirationDateTime), "The specified expiration date is invalid. Should be greater than today and within 6 months");

            await new SynchronizationContextRemover();

            string responseString = null;
            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(p => p.Url);
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, subscription.Resource);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }

                    string requestUrl = identifierUrl + "/" + SubscriptionsUrlPart;

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    request.Content = new StringContent(JsonConvert.SerializeObject(subscription),
                        Encoding.UTF8, "application/json");

                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        responseString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return JsonConvert.DeserializeObject<WebhookSubscription>(responseString);
        }

        /// <summary>
        /// Add a Webhook subscription to a SharePoint resource
        /// </summary>
        /// <param name="webUrl">Url of the site holding the list</param>
        /// <param name="resourceType">The type of Hookable SharePoint resource</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <param name="context">ClientContext instance to use for authentication</param>
        /// <param name="resourceId">The Unique Identifier of the resource</param>
        /// <param name="notificationUrl">The Webhook endpoint URL</param>
        /// <param name="clientState">The client state to use in the Webhook subscription</param>
        /// <param name="validityInMonths">The validity of the subscriptions in months</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when expiration date is out of valid range.</exception>
        /// <returns>The added subscription object</returns>
        public static async Task<WebhookSubscription> AddWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string accessToken, ClientContext context, string resourceId, string notificationUrl,
            string clientState = null, int validityInMonths = MaximumValidityInMonths)
        {
            await new SynchronizationContextRemover();

            // If validity in months is the Maximum, use the effective max allowed DateTime instead
            DateTime expirationDateTime = validityInMonths == MaximumValidityInMonths
                ? MaxExpirationDateTime
                : DateTime.UtcNow.AddMonths(validityInMonths);

            var subscription = new WebhookSubscription()
            {
                Resource = resourceId,
                NotificationUrl = notificationUrl,
                ExpirationDateTime = expirationDateTime,
                ClientState = clientState
            };

            return await AddWebhookSubscriptionAsync(webUrl, resourceType, accessToken, context, subscription);
        }

        /// <summary>
        /// Updates the expiration datetime (and notification URL) of an existing SharePoint web hook
        /// </summary>
        /// <param name="webUrl">Url of the site holding resource with the webhook</param>
        /// <param name="resourceType">The type of Hookable SharePoint resource</param>
        /// <param name="resourceId">Id of the resource (e.g. list id)</param>
        /// <param name="subscriptionId">Id of the web hook subscription that we need to delete</param>
        /// <param name="webHookEndPoint">Url of the web hook service endpoint (the one that will be called during an event)</param>
        /// <param name="expirationDateTime">New web hook expiration date</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <param name="context">ClientContext instance to use for authentication</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when expiration date is out of valid range.</exception>
        /// <returns>true if succesful, exception in case something went wrong</returns>
        public static async Task<bool> UpdateWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string resourceId, string subscriptionId,
            string webHookEndPoint, DateTime expirationDateTime, string accessToken, ClientContext context)
        {
            if (!ValidateExpirationDateTime(expirationDateTime))
                throw new ArgumentOutOfRangeException(nameof(expirationDateTime), "The specified expiration date is invalid. Should be greater than today and within 6 months");

            await new SynchronizationContextRemover();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(p => p.Url);
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }
                    
                    string requestUrl = string.Format("{0}/{1}('{2}')", identifierUrl, SubscriptionsUrlPart, subscriptionId);
                    HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), requestUrl);
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    WebhookSubscription webhookSubscription;

                    if (string.IsNullOrEmpty(webHookEndPoint))
                    {
                        webhookSubscription = new WebhookSubscription()
                        {
                            ExpirationDateTime = expirationDateTime
                        };
                    }
                    else
                    {
                        webhookSubscription = new WebhookSubscription()
                        {
                            NotificationUrl = webHookEndPoint,
                            ExpirationDateTime = expirationDateTime
                        };
                    }

                    request.Content = new StringContent(JsonConvert.SerializeObject(webhookSubscription),
                        Encoding.UTF8, "application/json");

                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.StatusCode != System.Net.HttpStatusCode.NoContent)
                    {
                        // oops...something went wrong, maybe the web hook does not exist?
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                    else
                    {
                        return true;
                    }
                }
            }
        }

        /// <summary>
        /// Deletes an existing SharePoint web hook
        /// </summary>
        /// <param name="webUrl">Url of the site holding resource with the webhook</param>
        /// <param name="resourceType">The type of Hookable SharePoint resource</param>
        /// <param name="resourceId">Id of the resource (e.g. list id)</param>
        /// <param name="subscriptionId">Id of the web hook subscription that we need to delete</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <param name="context">ClientContext instance to use for authentication</param>
        /// <returns>true if succesful, exception in case something went wrong</returns>
        public static async Task<bool> RemoveWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string resourceId, string subscriptionId, string accessToken, ClientContext context)
        {
            await new SynchronizationContextRemover();

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(p => p.Url);
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }

                    string requestUrl = string.Format("{0}/{1}('{2}')", identifierUrl, SubscriptionsUrlPart, subscriptionId);

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, requestUrl);
                    request.Headers.Add("X-RequestDigest", await context.GetRequestDigest());
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));


                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.StatusCode != System.Net.HttpStatusCode.NoContent)
                    {
                        // oops...something went wrong, maybe the web hook does not exist?
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                    else
                    {
                        return true;
                    }
                }
            }
        }

        /// <summary>
        /// Get all webhook subscriptions on a given SharePoint resource
        /// </summary>
        /// <param name="webUrl">Url of the site holding the resource</param>
        /// <param name="resourceType">The type of SharePoint resource</param>
        /// <param name="resourceId">The Unique Identifier of the resource</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <param name="context">ClientContext instance to use for authentication</param>
        /// <returns>Collection of <see cref="WebhookSubscription"/> instances, one per returned web hook</returns>
        public static async Task<ResponseModel<WebhookSubscription>> GetWebhooksSubscriptionsAsync(string webUrl, WebHookResourceType resourceType, string resourceId, string accessToken, ClientContext context)
        {
            await new SynchronizationContextRemover();

            string responseString = null;

            using (var handler = new HttpClientHandler())
            {
                context.Web.EnsureProperty(p => p.Url);
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(context);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }

                    string requestUrl = identifierUrl + "/" + SubscriptionsUrlPart;

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    if (accessToken != null)
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                    if (response.IsSuccessStatusCode)
                    {
                        responseString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        // oops...something went wrong
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            return JsonConvert.DeserializeObject<ResponseModel<WebhookSubscription>>(responseString);
        }

        /// <summary>
        /// Get the proper identifier Url according to the resource type. (No great utility currently with the support for lists only, but will be later with the support for other resources)
        /// </summary>
        /// <param name="resourceType">The type of resource</param>
        /// <param name="webUrl">The URL of the SharePoint web</param>
        /// <param name="id">The id part of the resource</param>
        /// <returns>The well forned resource identifier URL</returns>
        private static string GetResourceIdentifier(WebHookResourceType resourceType, string webUrl, string id)
        {
            switch (resourceType)
            {
                case WebHookResourceType.List:
                    return string.Format(ListIdentifierFormat, webUrl, id);
                default:
                    return null;
            }
        }

        /// <summary>
        /// Checks whether the specified expiration datetime is within a valid period
        /// </summary>
        /// <param name="expirationDateTime">The datetime value to validate</param>
        /// <returns><c>true</c> if valid, <c>false</c> otherwise</returns>
        private static bool ValidateExpirationDateTime(DateTime expirationDateTime)
        {
            DateTime utcDateToValidate = expirationDateTime.ToUniversalTime();
            DateTime utcNow = DateTime.UtcNow;

            return utcDateToValidate > utcNow
                && utcDateToValidate <= MaxExpirationDateTime;
        }
    }
}
#endif