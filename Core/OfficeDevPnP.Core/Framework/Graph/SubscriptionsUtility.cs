using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Graph
{
    /// <summary>
    /// Class that deals with Microsoft Graph Subscriptions
    /// </summary>
    public static class SubscriptionsUtility
    {
        /// <summary>
        /// Returns the subscription with the provided subscriptionId from Microsoft Graph
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="subscriptionId">The unique identifier of the subscription to return from Microsoft Graph</param>        
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <returns>Subscription object</returns>
        public static Model.Subscription GetSubscription(string accessToken, Guid subscriptionId, int startIndex = 0, int endIndex = 999, int retryCount = 10, int delay = 500)
        {
            try
            { 
                // Use a synchronous model to invoke the asynchronous process
                var result = Task.Run(async () =>
                {
                    var graphClient = GraphUtility.CreateGraphClient(accessToken, retryCount, delay);

                    var subscription = await graphClient.Subscriptions[subscriptionId.ToString()]
                        .Request()
                        .GetAsync();

                    var subscriptionModel = MapGraphEntityToModel(subscription);
                    return subscriptionModel;
                }).GetAwaiter().GetResult();

                return result;
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns all the active Microsoft Graph subscriptions
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <returns>List with Subscription objects</returns>
        public static List<Model.Subscription> ListSubscriptions(string accessToken, int startIndex = 0, int endIndex = 999, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            List<Model.Subscription> result = null;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    List<Model.Subscription> subscriptions = new List<Model.Subscription>();

                    var graphClient = GraphUtility.CreateGraphClient(accessToken, retryCount, delay);

                    var pagedSubscriptions = await graphClient.Subscriptions
                        .Request()
                        .GetAsync();

                    int pageCount = 0;
                    int currentIndex = 0;

                    while (true)
                    {
                        pageCount++;

                        foreach (var s in pagedSubscriptions)
                        {
                            currentIndex++;

                            if (currentIndex >= startIndex)
                            {
                                var subscription = MapGraphEntityToModel(s);
                                subscriptions.Add(subscription);
                            }
                        }

                        if (pagedSubscriptions.NextPageRequest != null && currentIndex < endIndex)
                        {
                            pagedSubscriptions = await pagedSubscriptions.NextPageRequest.GetAsync();
                        }
                        else
                        {
                            break;
                        }
                    }

                    return subscriptions;
                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return result;
        }

        /// <summary>
        /// Creates a new Microsoft Graph Subscription
        /// </summary>
        /// <param name="changeType">The event(s) the subscription should trigger on</param>
        /// <param name="notificationUrl">The URL that should be called when an event matching this subscription occurs</param>
        /// <param name="resource">The resource to monitor for changes. See https://docs.microsoft.com/graph/api/subscription-post-subscriptions#permissions for the list with supported options.</param>
        /// <param name="expirationDateTime">The datetime defining how long this subscription should stay alive before which it needs to get extended to stay alive. See https://docs.microsoft.com/graph/api/resources/subscription#maximum-length-of-subscription-per-resource-type for the supported maximum lifetime of the subscriber endpoints.</param>
        /// <param name="clientState">Specifies the value of the clientState property sent by the service in each notification. The maximum length is 128 characters. The client can check that the notification came from the service by comparing the value of the clientState property sent with the subscription with the value of the clientState property received with each notification.</param>
        /// <param name="latestSupportedTlsVersion">Specifies the latest version of Transport Layer Security (TLS) that the notification endpoint, specified by <paramref name="notificationUrl"/>, supports. If not provided, TLS 1.2 will be assumed.</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <returns>The just created Microsoft Graph subscription</returns>
        public static Model.Subscription CreateSubscription(Enums.GraphSubscriptionChangeType changeType, string notificationUrl, string resource, DateTimeOffset expirationDateTime, string clientState,
                                                            string accessToken, Enums.GraphSubscriptionTlsVersion latestSupportedTlsVersion = Enums.GraphSubscriptionTlsVersion.v1_2, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(notificationUrl))
            {
                throw new ArgumentNullException(nameof(notificationUrl));
            }

            if (String.IsNullOrEmpty(resource))
            {
                throw new ArgumentNullException(nameof(resource));
            }

            Model.Subscription result = null;

            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    var graphClient = GraphUtility.CreateGraphClient(accessToken, retryCount, delay);

                    // Prepare the subscription resource object
                    var newSubscription = new Subscription
                    {
                        ChangeType = changeType.ToString().Replace(" ", ""),
                        NotificationUrl = notificationUrl,
                        Resource = resource,
                        ExpirationDateTime = expirationDateTime,
                        ClientState = clientState
                    };

                    var subscription = await graphClient.Subscriptions
                                                        .Request()
                                                        .AddAsync(newSubscription);

                    if(subscription == null)
                    {
                        return null;
                    }

                    var subscriptionModel = MapGraphEntityToModel(subscription);
                    return subscriptionModel;

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return result;
        }

        /// <summary>
        /// Updates an existing Microsoft Graph Subscription
        /// </summary>
        /// <param name="subscriptionId">Unique identifier of the Microsoft Graph subscription</param>
        /// <param name="expirationDateTime">The datetime defining how long this subscription should stay alive before which it needs to get extended to stay alive. See https://docs.microsoft.com/graph/api/resources/subscription#maximum-length-of-subscription-per-resource-type for the supported maximum lifetime of the subscriber endpoints.</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <returns>The just updated Microsoft Graph subscription</returns>
        public static Model.Subscription UpdateSubscription(string subscriptionId, DateTimeOffset expirationDateTime, 
                                                            string accessToken, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(subscriptionId))
            {
                throw new ArgumentNullException(nameof(subscriptionId));
            }

            Model.Subscription result = null;

            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    var graphClient = GraphUtility.CreateGraphClient(accessToken, retryCount, delay);

                    // Prepare the subscription resource object
                    var updatedSubscription = new Subscription
                    {
                        ExpirationDateTime = expirationDateTime
                    };

                    var subscription = await graphClient.Subscriptions[subscriptionId]
                                                        .Request()
                                                        .UpdateAsync(updatedSubscription);

                    if (subscription == null)
                    {
                        return null;
                    }

                    var subscriptionModel = MapGraphEntityToModel(subscription);
                    return subscriptionModel;

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return result;
        }

        /// <summary>
        /// Deletes an existing Microsoft Graph Subscription
        /// </summary>
        /// <param name="subscriptionId">Unique identifier of the Microsoft Graph subscription</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        public static void DeleteSubscription(string subscriptionId,
                                              string accessToken, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(subscriptionId))
            {
                throw new ArgumentNullException(nameof(subscriptionId));
            }

            try
            {
                // Use a synchronous model to invoke the asynchronous process
                Task.Run(async () =>
                {
                    var graphClient = GraphUtility.CreateGraphClient(accessToken, retryCount, delay);

                    await graphClient.Subscriptions[subscriptionId]
                                     .Request()
                                     .DeleteAsync();

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Maps an entity returned by Microsoft Graph to its equivallent Model maintained within this library
        /// </summary>
        /// <param name="subscription">Microsoft Graph Subscription entity</param>
        /// <returns>Subscription Model</returns>
        private static Model.Subscription MapGraphEntityToModel(Subscription subscription)
        {
            var subscriptionModel = new Model.Subscription
            {
                Id = subscription.Id,
                ChangeType = subscription.ChangeType.Split(',').Select(ct => (Enums.GraphSubscriptionChangeType)Enum.Parse(typeof(Enums.GraphSubscriptionChangeType), ct, true)).Aggregate((prev, next) => prev | next),
                NotificationUrl = subscription.NotificationUrl,
                Resource = subscription.Resource,
                ExpirationDateTime = subscription.ExpirationDateTime,
                ClientState = subscription.ClientState
            };
            return subscriptionModel;
        }
    }
}
