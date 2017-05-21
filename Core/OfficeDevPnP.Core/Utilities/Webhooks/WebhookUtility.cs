using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities.Webhooks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    /// <summary>
    /// The list of Hookable Resource T^ypes
    /// </summary>
    public enum EHookableResourceType
    {
        List,
        // TODO Implement with upcoming support of other types
        //Site,
        //...
    }

    /// <summary>
    /// Class containing utility methods to manage Webhook on a SharePoint resource
    /// Adapted from https://github.com/SharePoint/sp-dev-samples/blob/master/Samples/WebHooks.List/SharePoint.WebHooks.Common/WebHookManager.cs
    /// 
    /// </summary>
    internal class WebhookUtility
    {

        private const string SubscriptionsUrlPart = "subscriptions";
        private const string ListIdentifierFormat = @"{0}/_api/web/lists('{1}')";
        // TODO Implement with upcoming support of other types
        //private const string WebIdentifierFormat = @"{0}/_api/web('{1}')";

        public static async Task<WebhookSubscription> AddWebhookSubscriptionAsync(string webUrl,
            EHookableResourceType resourceType,
        string accessToken, WebhookSubscription subscription)
        {
            string responseString = null;
            using (var httpClient = new HttpClient())
            {
                string identifierUrl = GetResourceIdentifier(resourceType, webUrl, subscription.Resource);
                if (string.IsNullOrEmpty(identifierUrl))
                    throw new Exception("Identifier of the resource cannot be determined");

                string requestUrl = identifierUrl + "/" + SubscriptionsUrlPart;

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                request.Content = new StringContent(JsonConvert.SerializeObject(subscription),
                    Encoding.UTF8, "application/json");

                HttpResponseMessage response = await httpClient.SendAsync(request);

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

            return JsonConvert.DeserializeObject<WebhookSubscription>(responseString);
        }

        public static async Task<WebhookSubscription> AddWebhookSubscriptionAsync(string webUrl,
            EHookableResourceType resourceType, string accessToken,
            string resourceId, string notificationUrl, string clientState = null, int validityInMonths = 3)
        {
            var subscription = new WebhookSubscription()
            {
                Resource = resourceId,
                NotificationUrl = notificationUrl,
                ExpirationDateTime = DateTime.Now.AddMonths(validityInMonths).ToUniversalTime(),
                ClientState = clientState
            };

            return await AddWebhookSubscriptionAsync(webUrl, resourceType, accessToken, subscription);
        }


        /// <summary>
        /// Deletes an existing SharePoint list web hook
        /// </summary>
        /// <param name="webUrl">Url of the site holding the list</param>
        /// <param name="resourceType">The type of Hookable SharePoint resource</param>
        /// <param name="resourceId">Id of the list</param>
        /// <param name="subscriptionId">Id of the web hook subscription that we need to delete</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <returns>true if succesful, exception in case something went wrong</returns>
        public static async Task<bool> DeleteWebhookSubscriptionAsync(string webUrl, EHookableResourceType resourceType,
            string resourceId, string subscriptionId, string accessToken)
        {
            using (var httpClient = new HttpClient())
            {
                string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                if (string.IsNullOrEmpty(identifierUrl))
                    throw new Exception("Identifier of the resource cannot be determined");

                string requestUrl = string.Format("{0}/{1}('{2}')", identifierUrl, SubscriptionsUrlPart, subscriptionId);

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, requestUrl);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                HttpResponseMessage response = await httpClient.SendAsync(request);

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

        /// <summary>
        /// Get all webhook subscriptions on a given SharePoint resource
        /// </summary>
        /// <param name="webUrl">Url of the site holding the list</param>
        /// <param name="resourceType">The type of Hookable SharePoint resource</param>
        /// <param name="resourceId">The Unique Identifier of the resource</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <returns>Collection of <see cref="WebhookSubscription"/> instances, one per returned web hook</returns>
        public static async Task<ResponseModel<WebhookSubscription>> GetWebhooksSubscriptionsAsync(string webUrl, EHookableResourceType resourceType, string resourceId, string accessToken)
        {
            string responseString = null;
            using (var httpClient = new HttpClient())
            {
                string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                if (string.IsNullOrEmpty(identifierUrl))
                    throw new Exception("Identifier of the resource cannot be determined");

                string requestUrl = identifierUrl + "/" + SubscriptionsUrlPart;

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                HttpResponseMessage response = await httpClient.SendAsync(request);

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

            return JsonConvert.DeserializeObject<ResponseModel<WebhookSubscription>>(responseString);
        }

        private static string GetResourceIdentifier(EHookableResourceType resourceType, string webUrl, string id)
        {
            switch (resourceType)
            {
                case EHookableResourceType.List:
                    return string.Format(ListIdentifierFormat, webUrl, id);
                //case EHookableResourceType.Site:
                default:
                    return null;
            }
        }
    }




}
