#if !ONPREMISES
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
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
        private const string ListIdentifierFormat = @"{0}/_api/web/lists('{1}')";

        /// <summary>
        /// Add a Webhook subscription to a SharePoint resource
        /// </summary>
        /// <param name="webUrl">Url of the site holding the list</param>
        /// <param name="resourceType">The type of Hookable SharePoint resource</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <param name="context">ClientContext instance to use for authentication</param>
        /// <param name="subscription">The Webhook subscription to add</param>
        /// <returns>The added subscription object</returns>
        internal static async Task<WebhookSubscription> AddWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string accessToken, ClientContext context, WebhookSubscription subscription)
        {
            string responseString = null;
            using (var handler = new HttpClientHandler())
            {
                if (String.IsNullOrEmpty(accessToken))
                {
                    context.Web.EnsureProperty(p => p.Url);
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, subscription.Resource);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }

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
        /// <returns>The added subscription object</returns>
        internal static async Task<WebhookSubscription> AddWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string accessToken, ClientContext context, string resourceId, string notificationUrl, 
            string clientState = null, int validityInMonths = 3)
        {
            var subscription = new WebhookSubscription()
            {
                Resource = resourceId,
                NotificationUrl = notificationUrl,
                ExpirationDateTime = DateTime.Now.AddMonths(validityInMonths).ToUniversalTime(),
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
        /// <returns>true if succesful, exception in case something went wrong</returns>
        internal static async Task<bool> UpdateWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string resourceId, string subscriptionId, string webHookEndPoint, DateTime expirationDateTime, string accessToken, ClientContext context)
        {
            using (var handler = new HttpClientHandler())
            {
                if (String.IsNullOrEmpty(accessToken))
                {
                    context.Web.EnsureProperty(p => p.Url);
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }

                    string requestUrl = string.Format("{0}/{1}('{2}')", identifierUrl, SubscriptionsUrlPart, subscriptionId);

                    HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), requestUrl);
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    request.Content = new StringContent(JsonConvert.SerializeObject(
                       new WebhookSubscription()
                       {
                           NotificationUrl = webHookEndPoint,
                           ExpirationDateTime = expirationDateTime.ToUniversalTime(),
                       }),
                       Encoding.UTF8, "application/json");

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
        internal static async Task<bool> RemoveWebhookSubscriptionAsync(string webUrl, WebHookResourceType resourceType, string resourceId, string subscriptionId, string accessToken, ClientContext context)
        {
            using (var handler = new HttpClientHandler())
            {
                if (String.IsNullOrEmpty(accessToken))
                {
                    context.Web.EnsureProperty(p => p.Url);
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }

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
        internal static async Task<ResponseModel<WebhookSubscription>> GetWebhooksSubscriptionsAsync(string webUrl, WebHookResourceType resourceType, string resourceId, string accessToken, ClientContext context)
        {
            string responseString = null;
            using (var handler = new HttpClientHandler())
            {
                if (String.IsNullOrEmpty(accessToken))
                {
                    context.Web.EnsureProperty(p => p.Url);
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new HttpClient(handler))
                {
                    string identifierUrl = GetResourceIdentifier(resourceType, webUrl, resourceId);
                    if (string.IsNullOrEmpty(identifierUrl))
                    {
                        throw new Exception("Identifier of the resource cannot be determined");
                    }

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
    }
}
#endif