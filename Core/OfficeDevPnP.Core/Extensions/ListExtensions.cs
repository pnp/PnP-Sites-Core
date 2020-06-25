using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Enums;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Utilities.Async;

#if !ONPREMISES
using OfficeDevPnP.Core.Utilities.Webhooks;
#endif

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that provides generic list creation and manipulation methods
    /// </summary>
    public static partial class ListExtensions
    {
        /// <summary>
        /// The common URL delimiters
        /// </summary>
        private static readonly char[] UrlDelimiters = { '\\', '/' };

        const string INDEXED_PROPERTY_KEY = "vti_indexedpropertykeys";

        #region Event Receivers

        /// <summary>
        /// Registers a remote event receiver
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="name">The name of the event receiver (needs to be unique among the event receivers registered on this list)</param>
        /// <param name="url">The URL of the remote WCF service that handles the event</param>
        /// <param name="eventReceiverType"></param>
        /// <param name="synchronization"></param>
        /// <param name="force">If True any event already registered with the same name will be removed first.</param>
        /// <returns>Returns an EventReceiverDefinition if succeeded. Returns null if failed.</returns>
        public static EventReceiverDefinition AddRemoteEventReceiver(this List list, string name, string url, EventReceiverType eventReceiverType, EventReceiverSynchronization synchronization, bool force)
        {
            return list.AddRemoteEventReceiver(name, url, eventReceiverType, synchronization, 1000, force);
        }

        /// <summary>
        /// Registers a remote event receiver
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="name">The name of the event receiver (needs to be unique among the event receivers registered on this list)</param>
        /// <param name="url">The URL of the remote WCF service that handles the event</param>
        /// <param name="eventReceiverType"></param>
        /// <param name="synchronization"></param>
        /// <param name="sequenceNumber"></param>
        /// <param name="force">If True any event already registered with the same name will be removed first.</param>
        /// <returns>Returns an EventReceiverDefinition if succeeded. Returns null if failed.</returns>
        public static EventReceiverDefinition AddRemoteEventReceiver(this List list, string name, string url, EventReceiverType eventReceiverType, EventReceiverSynchronization synchronization, int sequenceNumber, bool force)
        {
            var query = from receiver
                in list.EventReceivers
                        where receiver.ReceiverName == name
                        select receiver;
            var receivers = list.Context.LoadQuery(query);
            list.Context.ExecuteQueryRetry();

            var eventReceiverDefinitions = receivers as EventReceiverDefinition[] ?? receivers.ToArray();
            var receiverExists = eventReceiverDefinitions.Any();
            if (receiverExists && force)
            {
                var receiver = eventReceiverDefinitions.First();
                receiver.DeleteObject();
                list.Context.ExecuteQueryRetry();
                receiverExists = false;
            }
            EventReceiverDefinition def = null;

            if (!receiverExists)
            {
                EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();
                receiver.EventType = eventReceiverType;
                receiver.ReceiverUrl = url;
                receiver.ReceiverName = name;
                receiver.SequenceNumber = sequenceNumber;
                receiver.Synchronization = synchronization;
                def = list.EventReceivers.Add(receiver);
                list.Context.Load(def);
                list.Context.ExecuteQueryRetry();
            }
            return def;
        }

        /// <summary>
        /// Returns an event receiver definition
        /// </summary>
        /// <param name="list">The target list</param>
        /// <param name="id">Id of the event receiver</param>
        /// <returns></returns>
        public static EventReceiverDefinition GetEventReceiverById(this List list, Guid id)
        {
            var query = from receiver
                in list.EventReceivers
                        where receiver.ReceiverId == id
                        select receiver;

            var receivers = list.Context.LoadQuery(query);
            list.Context.ExecuteQueryRetry();
            var eventReceiverDefinitions = receivers as EventReceiverDefinition[] ?? receivers.ToArray();
            if (eventReceiverDefinitions.Any())
            {
                return eventReceiverDefinitions.FirstOrDefault();
            }
            return null;
        }

        /// <summary>
        /// Returns an event receiver definition
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="name">Name of the event receiver</param>
        /// <returns></returns>
        public static EventReceiverDefinition GetEventReceiverByName(this List list, string name)
        {
            var query = from receiver
                in list.EventReceivers
                        where receiver.ReceiverName == name
                        select receiver;

            var receivers = list.Context.LoadQuery(query);
            list.Context.ExecuteQueryRetry();
            var eventReceiverDefinitions = receivers as EventReceiverDefinition[] ?? receivers.ToArray();
            if (eventReceiverDefinitions.Any())
            {
                return eventReceiverDefinitions.FirstOrDefault();
            }
            return null;
        }

        #endregion

        #region Webhooks
#if !ONPREMISES
        /// <summary>
        /// Add the a Webhook subscription to a list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to add a Webhook subscription to</param>
        /// <param name="notificationUrl">The Webhook endpoint URL</param>
        /// <param name="expirationDate">The expiration date of the subscription</param>
        /// <param name="clientState">The client state to use in the Webhook subscription</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns>The added subscription object</returns>
        public static WebhookSubscription AddWebhookSubscription(this List list, string notificationUrl, DateTime expirationDate, string clientState = null, string accessToken = null)
        {
            // Get the access from the client context if not specified.
            accessToken = accessToken ?? list.Context.GetAccessToken();

            // Ensure the list Id is known
            Guid listId = list.EnsureProperty(l => l.Id);

            try
            {
                return WebhookUtility.AddWebhookSubscriptionAsync(list.Context.Url,
                               WebHookResourceType.List, accessToken, list.Context as ClientContext, new WebhookSubscription()
                               {
                                   Resource = listId.ToString(),
                                   ExpirationDateTime = expirationDate,
                                   NotificationUrl = notificationUrl,
                                   ClientState = clientState
                               }).Result;
            }
            catch (AggregateException ex)
            {
                // Rethrow the inner exception of the AggregateException thrown by the async method
                throw ex.InnerException ?? ex;
            }
        }

        /// <summary>
        /// Add the a Webhook subscription to a list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to add a Webhook subscription to</param>
        /// <param name="notificationUrl">The Webhook endpoint URL</param>
        /// <param name="validityInMonths">The validity of the subscriptions in months</param>
        /// <param name="clientState">The client state to use in the Webhook subscription</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns>The added subscription object</returns>
        public static WebhookSubscription AddWebhookSubscription(this List list, string notificationUrl, int validityInMonths = 6, string clientState = null, string accessToken = null)
        {
            // Get the access from the client context if not specified.
            accessToken = accessToken ?? list.Context.GetAccessToken();

            // Ensure the list Id is known
            Guid listId = list.EnsureProperty(l => l.Id);

            try
            {
                return WebhookUtility.AddWebhookSubscriptionAsync(list.Context.Url, WebHookResourceType.List, accessToken, list.Context as ClientContext, listId.ToString(), notificationUrl, clientState, validityInMonths).Result;
            }
            catch (AggregateException ex)
            {
                // Rethrow the inner exception of the AggregateException thrown by the async method
                throw ex.InnerException ?? ex;
            }
        }

        /// <summary>
        /// Updates a Webhook subscription from the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to remove the Webhook subscription from</param>
        /// <param name="subscriptionId">The id of the subscription to remove</param>
        /// <param name="webHookEndPoint">Url of the web hook service endpoint (the one that will be called during an event)</param>
        /// <param name="expirationDateTime">New web hook expiration date</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns><c>true</c> if the update succeeded, <c>false</c> otherwise</returns>
        public static bool UpdateWebhookSubscription(this List list, string subscriptionId, string webHookEndPoint, DateTime expirationDateTime, string accessToken = null)
        {
            // Get the access from the client context if not specified.
            accessToken = accessToken ?? list.Context.GetAccessToken();

            // Ensure the list Id is known
            Guid listId = list.EnsureProperty(l => l.Id);

            try
            {
                return WebhookUtility.UpdateWebhookSubscriptionAsync(list.Context.Url, WebHookResourceType.List, listId.ToString(), subscriptionId, webHookEndPoint, expirationDateTime, accessToken, list.Context as ClientContext).Result;
            }
            catch (AggregateException ex)
            {
                // Rethrow the inner exception of the AggregateException thrown by the async method
                throw ex.InnerException ?? ex;
            }
        }

        /// <summary>
        /// Updates a Webhook subscription from the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to remove the Webhook subscription from</param>
        /// <param name="subscriptionId">The id of the subscription to remove</param>
        /// <param name="expirationDateTime">New web hook expiration date</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns><c>true</c> if the update succeeded, <c>false</c> otherwise</returns>
        public static bool UpdateWebhookSubscription(this List list, Guid subscriptionId, DateTime expirationDateTime, string accessToken = null)
        {
            return UpdateWebhookSubscription(list, subscriptionId.ToString(), null, expirationDateTime, accessToken);
        }

        /// <summary>
        /// Updates a Webhook subscription from the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to remove the Webhook subscription from</param>
        /// <param name="subscriptionId">The id of the subscription to remove</param>
        /// <param name="webHookEndPoint">Url of the web hook service endpoint (the one that will be called during an event)</param>
        /// <param name="expirationDateTime">New web hook expiration date</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns><c>true</c> if the update succeeded, <c>false</c> otherwise</returns>
        public static bool UpdateWebhookSubscription(this List list, Guid subscriptionId, string webHookEndPoint, DateTime expirationDateTime, string accessToken = null)
        {
            return UpdateWebhookSubscription(list, subscriptionId.ToString(), webHookEndPoint, expirationDateTime, accessToken);
        }

        /// <summary>
        /// Updates a Webhook subscription from the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to remove the Webhook subscription from</param>
        /// <param name="subscription">The subscription to update</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns><c>true</c> if the update succeeded, <c>false</c> otherwise</returns>
        public static bool UpdateWebhookSubscription(this List list, WebhookSubscription subscription, string accessToken = null)
        {
            return UpdateWebhookSubscription(list, subscription.Id, subscription.NotificationUrl, subscription.ExpirationDateTime, accessToken);
        }

        /// <summary>
        /// Remove a Webhook subscription from the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to remove the Webhook subscription from</param>
        /// <param name="subscriptionId">The id of the subscription to remove</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns><c>true</c> if the removal succeeded, <c>false</c> otherwise</returns>
        public static bool RemoveWebhookSubscription(this List list, string subscriptionId, string accessToken = null)
        {
            // Get the access from the client context if not specified.
            accessToken = accessToken ?? list.Context.GetAccessToken();

            // Ensure the list Id is known
            Guid listId = list.EnsureProperty(l => l.Id);

            try
            {
                return WebhookUtility.RemoveWebhookSubscriptionAsync(list.Context.Url, WebHookResourceType.List, listId.ToString(), subscriptionId, accessToken, list.Context as ClientContext).Result;
            }
            catch (AggregateException ex)
            {
                // Rethrow the inner exception of the AggregateException thrown by the async method
                throw ex.InnerException ?? ex;
            }
        }

        /// <summary>
        /// Remove a Webhook subscription from the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to remove the Webhook subscription from</param>
        /// <param name="subscriptionId">The id of the subscription to remove</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns><c>true</c> if the removal succeeded, <c>false</c> otherwise</returns>
        public static bool RemoveWebhookSubscription(this List list, Guid subscriptionId, string accessToken = null)
        {
            return RemoveWebhookSubscription(list, subscriptionId.ToString(), accessToken);
        }

        /// <summary>
        /// Remove a Webhook subscription from the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to remove the Webhook subscription from</param>
        /// <param name="subscription">The subscription to remove</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns><c>true</c> if the removal succeeded, <c>false</c> otherwise</returns>
        public static bool RemoveWebhookSubscription(this List list, WebhookSubscription subscription, string accessToken = null)
        {
            return RemoveWebhookSubscription(list, subscription.Id, accessToken);
        }

        /// <summary>
        /// Get all the existing Webhooks subscriptions of the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to get the subscriptions of</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns>The collection of Webhooks subscriptions of the list</returns>
        public static IList<WebhookSubscription> GetWebhookSubscriptions(this List list, string accessToken = null)
        {
            // Get the access from the client context if not specified.
            accessToken = accessToken ?? list.Context.GetAccessToken();

            // Ensure the list Id is known
            Guid listId = list.EnsureProperty(l => l.Id);

            try
            {
                return WebhookUtility.GetWebhooksSubscriptionsAsync(list.Context.Url, WebHookResourceType.List, listId.ToString(), accessToken, list.Context as ClientContext).Result.Value;
            }
            catch (AggregateException ex)
            {
                // Rethrow the inner exception of the AggregateException thrown by the async method
                throw ex.InnerException ?? ex;
            }
        }

        /// <summary>
        /// Async get all the existing Webhooks subscriptions of the list
        /// Note: If the access token is not specified, it will cost a dummy request to retrieve it
        /// </summary>
        /// <param name="list">The list to get the subscriptions of</param>
        /// <param name="accessToken">(optional) The access token to SharePoint</param>
        /// <returns>The collection of Webhooks subscriptions of the list</returns>
        public static async Task<IList<WebhookSubscription>> GetWebhookSubscriptionsAsync(this List list, string accessToken = null)
        {
            await new SynchronizationContextRemover();

            // Get the access from the client context if not specified.
            accessToken = accessToken ?? list.Context.GetAccessToken();

            // Ensure the list Id is known
            Guid listId = list.EnsureProperty(l => l.Id);

            try
            {
                ResponseModel<WebhookSubscription> webHookSubscriptionResponse = await WebhookUtility.GetWebhooksSubscriptionsAsync(list.Context.Url, WebHookResourceType.List, listId.ToString(), accessToken, list.Context as ClientContext);
                return webHookSubscriptionResponse.Value;
            }
            catch (AggregateException ex)
            {
                // Rethrow the inner exception of the AggregateException thrown by the async method
                throw ex.InnerException ?? ex;
            }
        }
#endif
        #endregion

        #region List Properties

        /// <summary>
        /// Sets a key/value pair in the list property bag
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">Integer value for the property bag entry</param>
        public static void SetPropertyBagValue(this List list, string key, int value)
        {
            SetPropertyBagValueInternal(list, key, value);
        }


        /// <summary>
        /// Sets a key/value pair in the list property bag
        /// </summary>
        /// <param name="list">List that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">String value for the property bag entry</param>
        public static void SetPropertyBagValue(this List list, string key, string value)
        {
            SetPropertyBagValueInternal(list, key, value);
        }

        /// <summary>
        /// Sets a key/value pair in the list property bag
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">Datetime value for the property bag entry</param>
        public static void SetPropertyBagValue(this List list, string key, DateTime value)
        {
            SetPropertyBagValueInternal(list, key, value);
        }

        /// <summary>
        /// Sets a key/value pair in the list property bag
        /// </summary>
        /// <param name="list">List that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">Value for the property bag entry</param>
        private static void SetPropertyBagValueInternal(List list, string key, object value)
        {
            var props = list.RootFolder.Properties;
            list.Context.Load(props);
            list.Context.ExecuteQueryRetry();

            props[key] = value;
            list.RootFolder.Update();
            list.Update();
            list.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Removes a property bag value from the property bag
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="key">The key to remove</param>
        public static void RemovePropertyBagValue(this List list, string key)
        {
            RemovePropertyBagValueInternal(list, key, true);
        }

        /// <summary>
        /// Removes a property bag value
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="key">They key to remove</param>
        /// <param name="checkIndexed"></param>
        private static void RemovePropertyBagValueInternal(List list, string key, bool checkIndexed)
        {
            // In order to remove a property from the property bag, remove it both from the Properties collection by setting it to null
            // -and- by removing it from the FieldValues collection. Bug in CSOM?
            list.RootFolder.Properties[key] = null;
            list.RootFolder.Properties.FieldValues.Remove(key);

            list.RootFolder.Update();

            list.Context.ExecuteQueryRetry();
            if (checkIndexed)
                RemoveIndexedPropertyBagKey(list, key); // Will only remove it if it exists as an indexed property
        }

        /// <summary>
        /// Get int typed property bag value. If does not contain, returns default value.
        /// </summary>
        /// <param name="list">List to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <param name="defaultValue">Default value of the property bag</param>
        /// <returns>Value of the property bag entry as integer</returns>
        public static int? GetPropertyBagValueInt(this List list, string key, int defaultValue)
        {
            object value = GetPropertyBagValueInternal(list, key);
            if (value != null)
            {
                return (int)value;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Get string typed property bag value. If does not contain, returns given default value.
        /// </summary>
        /// <param name="list">List to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <param name="defaultValue">Default value of the property bag</param>
        /// <returns>Value of the property bag entry as string</returns>
        public static string GetPropertyBagValueString(this List list, string key, string defaultValue)
        {
            object value = GetPropertyBagValueInternal(list, key);
            if (value != null)
            {
                return (string)value;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Get DateTime typed property bag value. If does not contain, returns default value.
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <param name="defaultValue"></param>
        /// <returns>Value of the property bag entry as integer</returns>
        public static DateTime? GetPropertyBagValueDateTime(this List list, string key, DateTime defaultValue)
        {
            object value = GetPropertyBagValueInternal(list, key);
            if (value != null)
            {
                return (DateTime)value;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Type independent implementation of the property gettter.
        /// </summary>
        /// <param name="list">List to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <returns>Value of the property bag entry</returns>
        private static object GetPropertyBagValueInternal(List list, string key)
        {
            var props = list.RootFolder.Properties;
            list.Context.Load(props);
            list.Context.ExecuteQueryRetry();
            if (props.FieldValues.ContainsKey(key))
            {
                return props.FieldValues[key];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Checks if the given property bag entry exists
        /// </summary>
        /// <param name="list">List to be processed</param>
        /// <param name="key">Key of the property bag entry to check</param>
        /// <returns>True if the entry exists, false otherwise</returns>
        public static bool PropertyBagContainsKey(this List list, string key)
        {
            var props = list.RootFolder.Properties;
            list.Context.Load(props);
            list.Context.ExecuteQueryRetry();
            if (props.FieldValues.ContainsKey(key))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Used to convert the list of property keys is required format for listing keys to be index
        /// </summary>
        /// <param name="keys">list of keys to set to be searchable</param>
        /// <returns>string formatted list of keys in proper format</returns>
        private static string GetEncodedValueForSearchIndexProperty(IEnumerable<string> keys)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (string current in keys)
            {
                stringBuilder.Append(Convert.ToBase64String(Encoding.Unicode.GetBytes(current)));
                stringBuilder.Append('|');
            }
            return stringBuilder.ToString();
        }

        /// <summary>
        /// Returns all keys in the property bag that have been marked for indexing
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <returns>all indexed property bag keys</returns>
        public static IEnumerable<string> GetIndexedPropertyBagKeys(this List list)
        {
            var keys = new List<string>();

            if (list.PropertyBagContainsKey(INDEXED_PROPERTY_KEY))
            {
                foreach (var key in list.GetPropertyBagValueString(INDEXED_PROPERTY_KEY, "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var bytes = Convert.FromBase64String(key);
                    keys.Add(Encoding.Unicode.GetString(bytes));
                }
            }

            return keys;
        }

        /// <summary>
        /// Marks a property bag key for indexing
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="key">The key to mark for indexing</param>
        /// <returns>Returns True if succeeded</returns>
        public static bool AddIndexedPropertyBagKey(this List list, string key)
        {
            var result = false;
            var keys = GetIndexedPropertyBagKeys(list).ToList();
            if (!keys.Contains(key))
            {
                keys.Add(key);
                list.SetPropertyBagValue(INDEXED_PROPERTY_KEY, GetEncodedValueForSearchIndexProperty(keys));
                result = true;
            }
            return result;
        }

        /// <summary>
        /// Unmarks a property bag key for indexing
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="key">The key to unmark for indexed. Case-sensitive</param>
        /// <returns>Returns True if succeeded</returns>
        public static bool RemoveIndexedPropertyBagKey(this List list, string key)
        {
            var result = false;
            var keys = GetIndexedPropertyBagKeys(list).ToList();
            if (key.Contains(key))
            {
                keys.Remove(key);
                if (keys.Any())
                {
                    list.SetPropertyBagValue(INDEXED_PROPERTY_KEY, GetEncodedValueForSearchIndexProperty(keys));
                }
                else
                {
                    RemovePropertyBagValueInternal(list, INDEXED_PROPERTY_KEY, false);
                }
                result = true;
            }
            return result;
        }


        #endregion

        /// <summary>
        /// Removes a content type from a list/library by name
        /// </summary>
        /// <param name="list">The list</param>
        /// <param name="contentTypeName">The content type name to remove from the list</param>
        /// <exception cref="System.ArgumentException">Thrown when contentTypeName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">contentTypeName is null</exception>
        public static void RemoveContentTypeByName(this List list, string contentTypeName)
        {
            if (string.IsNullOrEmpty(contentTypeName))
            {
                throw (contentTypeName == null)
                    ? new ArgumentNullException(nameof(contentTypeName))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(contentTypeName));
            }

            var cts = list.ContentTypes;
            list.Context.Load(cts);

            var results = list.Context.LoadQuery<ContentType>(cts.Where(item => item.Name == contentTypeName));
            list.Context.ExecuteQueryRetry();

            var ct = results.FirstOrDefault();
            if (ct != null)
            {
                ct.DeleteObject();
                list.Update();
                list.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Adds a document library to a web. Execute Query is called during this implementation
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listName">Name of the library</param>
        /// <param name="enableVersioning">Enable versioning on the list</param>
        /// <param name="urlPath">Path of the url</param>
        /// <exception cref="System.ArgumentException">Thrown when listName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">listName is null</exception>
        public static List CreateDocumentLibrary(this Web web, string listName, bool enableVersioning = false, string urlPath = "")
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw (listName == null)
                    ? new ArgumentNullException(nameof(listName))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(listName));
            }
            // Call actual implementation
            return CreateListInternal(web, null, (int)ListTemplateType.DocumentLibrary, listName, enableVersioning, urlPath: urlPath);
        }

        /// <summary>
        /// Checks if list exists on the particular site based on the list Title property.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listTitle">Title of the list to be checked.</param>
        /// <exception cref="System.ArgumentException">Thrown when listTitle is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">listTitle is null</exception>
        /// <returns>True if the list exists</returns>
        public static bool ListExists(this Web web, string listTitle)
        {
            if (string.IsNullOrEmpty(listTitle))
            {
                throw (listTitle == null)
                    ? new ArgumentNullException(nameof(listTitle))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(listTitle));
            }

            var lists = web.Lists;
            var results = web.Context.LoadQuery(lists.Where(list => list.Title == listTitle));
            web.Context.ExecuteQueryRetry();
            //var existingList = results.FirstOrDefault();

            return results.Any();
        }

        /// <summary>
        /// Checks if list exists on the particular site based on the list's site relative path.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="siteRelativeUrlPath">Site relative path of the list</param>
        /// <returns>True if the list exists</returns>
        public static bool ListExists(this Web web, Uri siteRelativeUrlPath)
        {
            if (siteRelativeUrlPath == null)
            {
                throw new ArgumentNullException(nameof(siteRelativeUrlPath));
            }

            web.EnsureProperty(w => w.ServerRelativeUrl);

            var listResult = web.GetList(UrlUtility.Combine(web.ServerRelativeUrl, siteRelativeUrlPath.ToString()));
            web.Context.ExecuteQueryRetry();

            return listResult != null;
        }

        /// <summary>
        /// Checks if list exists on the particular site based on the list id property.
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="id">The id of the list to be checked.</param>
        /// <exception cref="System.ArgumentException">Thrown when listTitle is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">listTitle is null</exception>
        /// <returns>True if the list exists</returns>
        public static bool ListExists(this Web web, Guid id)
        {
            if (id == Guid.Empty)
            {
                throw new ArgumentException(nameof(id));
            }

            var lists = web.Lists;
            var results = web.Context.LoadQuery(lists.Where(list => list.Id == id));
            web.Context.ExecuteQueryRetry();
            var existingList = results.FirstOrDefault();

            if (existingList != null)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Adds a default list to a site
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listType">Built in list template type</param>
        /// <param name="listName">Name of the list</param>
        /// <param name="enableVersioning">Enable versioning on the list</param>
        /// <param name="updateAndExecuteQuery">(Optional) Perform list update and executequery, defaults to true</param>
        /// <param name="urlPath">(Optional) URL to use for the list</param>
        /// <param name="enableContentTypes">(Optional) Enable content type management</param>
        /// <param name="hidden">(Optional) Hide the list from the SharePoint UI</param>
        /// <returns>The newly created list</returns>
        public static List CreateList(this Web web, ListTemplateType listType, string listName, bool enableVersioning, bool updateAndExecuteQuery = true, string urlPath = "", bool enableContentTypes = false, bool hidden = false)
        {
            return CreateListInternal(web, null, (int)listType, listName, enableVersioning, updateAndExecuteQuery, urlPath, enableContentTypes, hidden);
        }

        /// <summary>
        /// Adds a custom list to a site
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="featureId">Feature that contains the list template</param>
        /// <param name="listType">Type ID of the list, within the feature</param>
        /// <param name="listName">Name of the list</param>
        /// <param name="enableVersioning">Enable versioning on the list</param>
        /// <param name="updateAndExecuteQuery">(Optional) Perform list update and executequery, defaults to true</param>
        /// <param name="urlPath">(Optional) URL to use for the list</param>
        /// <param name="enableContentTypes">(Optional) Enable content type management</param>
        /// <returns>The newly created list</returns>
        public static List CreateList(this Web web, Guid featureId, int listType, string listName, bool enableVersioning, bool updateAndExecuteQuery = true, string urlPath = "", bool enableContentTypes = false)
        {
            return CreateListInternal(web, featureId, listType, listName, enableVersioning, updateAndExecuteQuery, urlPath, enableContentTypes);
        }

        private static List CreateListInternal(this Web web, Guid? templateFeatureId, int templateType, string listName, bool enableVersioning, bool updateAndExecuteQuery = true, string urlPath = "", bool enableContentTypes = false, bool hidden = false)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.ListExtensions_CreateList0Template12, listName, templateType, templateFeatureId.HasValue ? " (feature " + templateFeatureId.Value.ToString() + ")" : "");

            ListCollection listCol = web.Lists;
            var lci = new ListCreationInformation
            {
                Title = listName,
                TemplateType = templateType
            };
            if (templateFeatureId.HasValue)
            {
                lci.TemplateFeatureId = templateFeatureId.Value;
            }
            if (!string.IsNullOrEmpty(urlPath))
            {
                lci.Url = urlPath;
            }

            List newList = listCol.Add(lci);

            if (enableVersioning)
            {
                newList.EnableVersioning = true;
                if (templateType == (int)ListTemplateType.DocumentLibrary)
                {
                    newList.EnableMinorVersions = true;
                }
            }
            if (enableContentTypes)
            {
                newList.ContentTypesEnabled = true;
            }
            if (hidden)
            {
                newList.Hidden = true;
            }
            if (updateAndExecuteQuery)
            {
                newList.Update();
                web.Context.Load(listCol);
                web.Context.ExecuteQueryRetry();
            }

            return newList;
        }

        /// <summary>
        /// Enable/disable versioning on a list
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listName">List to operate on</param>
        /// <param name="enableVersioning">True to enable versioning, false to disable</param>
        /// <param name="enableMinorVersioning">Enable/Disable minor versioning</param>
        /// <param name="updateAndExecuteQuery">Perform list update and executequery, defaults to true</param>
        /// <exception cref="System.ArgumentException">Thrown when listName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">listName is null</exception>
        public static void UpdateListVersioning(this Web web, string listName, bool enableVersioning, bool enableMinorVersioning = true, bool updateAndExecuteQuery = true)
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw (listName == null)
                    ? new ArgumentNullException(nameof(listName))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(listName));
            }

            List listToUpdate = web.Lists.GetByTitle(listName);
            listToUpdate.EnableVersioning = enableVersioning;
            listToUpdate.EnableMinorVersions = enableMinorVersioning;

            if (updateAndExecuteQuery)
            {
                listToUpdate.Update();
                web.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Enable/disable versioning on a list
        /// </summary>
        /// <param name="list">List to be processed</param>
        /// <param name="enableVersioning">True to enable versioning, false to disable</param>
        /// <param name="enableMinorVersioning">Enable/Disable minor versioning</param>
        /// <param name="updateAndExecuteQuery">Perform list update and executequery, defaults to true</param>
        public static void UpdateListVersioning(this List list, bool enableVersioning, bool enableMinorVersioning = true, bool updateAndExecuteQuery = true)
        {
            list.EnableVersioning = enableVersioning;
            list.EnableMinorVersions = enableMinorVersioning;

            if (updateAndExecuteQuery)
            {
                list.Update();
                list.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Sets the default value for a managed metadata column in the specified list. This operation will not change existing items in the list.
        /// </summary>
        /// <param name="web">Extension web</param>
        /// <param name="termName">Name of a specific term which should be set as the default on the managed metadata field</param>
        /// <param name="listName">Name of list which contains the managed metadata field of which the default needs to be set</param>
        /// <param name="fieldInternalName">Internal name of the managed metadata field for which the default needs to be set</param>
        /// <param name="groupGuid">TermGroup Guid of the Term Group which contains the managed metadata item which should be set as the default</param>
        /// <param name="termSetGuid">TermSet Guid of the Term Set which contains the managed metadata item which should be set as the default</param>
        /// <param name="systemUpdate">If set to true, will do a system udpate to the item. Default value is false.</param>
        public static void UpdateTaxonomyFieldDefaultValue(this Web web, string termName, string listName, string fieldInternalName, Guid groupGuid, Guid termSetGuid, bool systemUpdate = false)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            var termGroup = termStore.GetGroup(groupGuid);
            var termSet = termGroup.TermSets.GetById(termSetGuid);
            var term = web.Context.LoadQuery(termSet.Terms.Where(t => t.Name == termName));

            web.Context.ExecuteQueryRetry();

            var foundTerm = term.First();

            web.UpdateTaxonomyFieldDefaultValue(listName, fieldInternalName, foundTerm, systemUpdate);
        }

        /// <summary>
        /// Sets the default value for a managed metadata column in the specified list. This operation will not change existing items in the list.
        /// </summary>
        /// <param name="web">Extension web</param>
        /// <param name="listName">Name of list which contains the managed metadata field of which the default needs to be set</param>
        /// <param name="fieldInternalName">Internal name of the managed metadata field for which the default needs to be set</param>
        /// <param name="termGuid">Term Guid of the Term which represents the managed metadata item which should be set as the default</param>
        /// <param name="systemUpdate">If set to true, will do a system udpate to the item. Default value is false.</param>
        public static void UpdateTaxonomyFieldDefaultValue(this Web web, string listName, string fieldInternalName, Guid termGuid, bool systemUpdate = false)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(web.Context);
            var foundTerm = taxonomySession.GetTerm(termGuid);

            web.UpdateTaxonomyFieldDefaultValue(listName, fieldInternalName, foundTerm, systemUpdate);

        }

        /// <summary>
        /// Sets the default value for a managed metadata column in the specified list. This operation will not change existing items in the list.
        /// </summary>
        /// <param name="web">Extension web</param>
        /// <param name="listName">Name of list which contains the managed metadata field of which the default needs to be set</param>
        /// <param name="fieldInternalName">Internal name of the managed metadata field for which the default needs to be set</param>
        /// <param name="term">Managed metadata Term which represents the managed metadata item which should be set as the default</param>
        /// <param name="systemUpdate">If set to true, will do a system udpate to the item. Default value is false.</param>
        public static void UpdateTaxonomyFieldDefaultValue(this Web web, string listName, string fieldInternalName, Term term, bool systemUpdate = false)
        {
            var list = web.GetListByTitle(listName);

            var fields = web.Context.LoadQuery(list.Fields.Where(f => f.InternalName == fieldInternalName));
            web.Context.ExecuteQueryRetry();

            var taxField = web.Context.CastTo<TaxonomyField>(fields.First());

            //The default value requires that the item is present in the TaxonomyHiddenList (which gives it it's WssId)
            //To solve this, we create a folder that we assign the value, which creates the listitem in the hidden list
            var item = list.AddItem(new ListItemCreationInformation()
            {
                UnderlyingObjectType = FileSystemObjectType.Folder,
                LeafName = string.Concat("Temporary_Folder_For_WssId_Creation_", DateTime.Now.ToFileTime().ToString())
            });

            item.SetTaxonomyFieldValue(taxField.Id, term.Name, term.Id, systemUpdate);

            web.Context.Load(item);
            web.Context.ExecuteQueryRetry();

            dynamic val = item[fieldInternalName];

            //The folder has now served it's purpose and can safely be removed
            item.DeleteObject();

            taxField.DefaultValue = string.Format("{0};#{1}|{2}", val.WssId, val.Label, val.TermGuid);
            taxField.Update();

            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Sets JS link customization for a list form
        /// </summary>
        /// <param name="list">SharePoint list</param>
        /// <param name="pageType">Type of form</param>
        /// <param name="jslink">JSLink to set to the form. Set to empty string to remove the set JSLink customization.
        /// Specify multiple values separated by pipe symbol. For e.g.: ~sitecollection/_catalogs/masterpage/jquery-2.1.0.min.js|~sitecollection/_catalogs/masterpage/custom.js
        /// </param>
        public static void SetJSLinkCustomizations(this List list, PageType pageType, string jslink)
        {
            // Get the list form to apply the JS link
            Form listForm = list.Forms.GetByPageType(pageType);
            list.Context.Load(listForm, nf => nf.ServerRelativeUrl);
            list.Context.ExecuteQueryRetry();

            var file = list.ParentWeb.GetFileByServerRelativeUrl(listForm.ServerRelativeUrl);
            SetJSLinkCustomizationsImplementation(list, file, jslink);
        }

        /// <summary>
        /// Sets JS link customization for a list view page
        /// </summary>
        /// <param name="list">SharePoint list</param>
        /// <param name="serverRelativeUrl">url of the view page</param>
        /// <param name="jslink">JSLink to set to the form. Set to empty string to remove the set JSLink customization.
        /// Specify multiple values separated by pipe symbol. For e.g.: ~sitecollection/_catalogs/masterpage/jquery-2.1.0.min.js|~sitecollection/_catalogs/masterpage/custom.js
        /// </param>
        public static void SetJSLinkCustomizations(this List list, string serverRelativeUrl, string jslink)
        {

            var file = list.ParentWeb.GetFileByServerRelativeUrl(serverRelativeUrl);
            SetJSLinkCustomizationsImplementation(list, file, jslink);
        }

        private static void SetJSLinkCustomizationsImplementation(List list, File file, string jslink)
        {
            var wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
            list.Context.Load(wpm.WebParts, wps => wps.Include(wp => wp.WebPart.Title, wp => wp.WebPart.Properties));
            list.Context.ExecuteQueryRetry();

            // Set the JS link for all web parts
            foreach (var wpd in wpm.WebParts)
            {
                var wp = wpd.WebPart;

                if (wp.Properties.FieldValues.Keys.Contains("JSLink"))
                {
                    wp.Properties["JSLink"] = jslink;
                    wpd.SaveWebPartChanges();

                    list.Context.ExecuteQueryRetry();
                }
            }
        }


#if !SP2013 && !SP2016
        /// <summary>
        /// Can be used to set translations for different cultures. 
        /// <see href="http://blogs.msdn.com/b/vesku/archive/2014/03/20/office365-multilingual-content-types-site-columns-and-site-other-elements.aspx"/>
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listTitle">Title of the list</param>
        /// <param name="cultureName">Culture name like en-us or fi-fi</param>
        /// <param name="titleResource">Localized Title string</param>
        /// <param name="descriptionResource">Localized Description string</param>
        /// <exception cref="System.ArgumentException">Thrown when listTitle, cultureName, titleResource, descriptionResource is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">listTitle, cultureName, titleResource, descriptionResource is null</exception>
        public static void SetLocalizationLabelsForList(this Web web, string listTitle, string cultureName, string titleResource, string descriptionResource)
        {
            if (string.IsNullOrEmpty(listTitle))
            {
                throw (listTitle == null)
                    ? new ArgumentNullException(nameof(listTitle))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(listTitle));
            }
            if (string.IsNullOrEmpty(cultureName))
            {
                throw (cultureName == null)
                    ? new ArgumentNullException(nameof(cultureName))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(cultureName));
            }
            if (string.IsNullOrEmpty(titleResource))
            {
                throw (titleResource == null)
                    ? new ArgumentNullException(nameof(titleResource))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(titleResource));
            }
            if (string.IsNullOrEmpty(descriptionResource))
            {
                throw (descriptionResource == null)
                    ? new ArgumentNullException(nameof(descriptionResource))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(descriptionResource));
            }

            List list = web.GetList(listTitle);
            SetLocalizationLabelsForList(list, cultureName, titleResource, descriptionResource);
        }

        /// <summary>
        /// Can be used to set translations for different cultures. 
        /// </summary>
        /// <example>
        ///     list.SetLocalizationForSiteLabels("fi-fi", "Name of the site in Finnish", "Description in Finnish");
        /// </example>
        /// <see href="http://blogs.msdn.com/b/vesku/archive/2014/03/20/office365-multilingual-content-types-site-columns-and-site-other-elements.aspx"/>
        /// <param name="list">List to be processed </param>
        /// <param name="cultureName">Culture name like en-us or fi-fi</param>
        /// <param name="titleResource">Localized Title string</param>
        /// <param name="descriptionResource">Localized Description string</param>
        public static void SetLocalizationLabelsForList(this List list, string cultureName, string titleResource, string descriptionResource)
        {
            list.TitleResource.SetValueForUICulture(cultureName, titleResource);
            list.DescriptionResource.SetValueForUICulture(cultureName, descriptionResource);
            list.Update();
            list.Context.ExecuteQueryRetry();
        }
#endif

        /// <summary>
        /// Returns the GUID id of a list
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listName">List to operate on</param>
        /// <exception cref="System.ArgumentException">Thrown when listName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">listName is null</exception>
        public static Guid GetListID(this Web web, string listName)
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw (listName == null)
                    ? new ArgumentNullException(nameof(listName))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(listName));
            }

            var listToQuery = web.Lists.GetByTitle(listName);
            web.Context.Load(listToQuery, l => l.Id);
            web.Context.ExecuteQueryRetry();

            return listToQuery.Id;
        }

        /// <summary>
        /// Get List by using Id
        /// </summary>
        /// <param name="web">The web containing the list</param>
        /// <param name="listId">The Id of the list</param>
        /// <param name="expressions">Additional list of lambda expressions of properties to load alike l => l.BaseType</param>
        /// <returns>Loaded list instance matching specified Id</returns>
        /// <exception cref="System.ArgumentException">Thrown when listId is an empty Guid</exception>
        /// <exception cref="System.ArgumentNullException">listId is null</exception>
        public static List GetListById(this Web web, Guid listId, params Expression<Func<List, object>>[] expressions)
        {
            var baseExpressions = new List<Expression<Func<List, object>>> { l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, l => l.Hidden, l => l.RootFolder };

            if (expressions != null && expressions.Any())
            {
                baseExpressions.AddRange(expressions);
            }

            if (listId == null)
            {
                throw new ArgumentNullException(nameof(listId));
            }

            if (listId == Guid.Empty)
            {
                throw new ArgumentException(nameof(listId));
            }

            var query = web.Lists.IncludeWithDefaultProperties(baseExpressions.ToArray());
            var lists = web.Context.LoadQuery(query.Where(l => l.Id == listId));

            web.Context.ExecuteQueryRetry();

            return lists.FirstOrDefault();
        }

        /// <summary>
        /// Get list by using Title
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="listTitle">Title of the list to return</param>
        /// <returns>Loaded list instance matching to title or null</returns>
        /// <exception cref="System.ArgumentException">Thrown when listTitle is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">listTitle is null</exception>
        /// <param name="expressions">Additional list of lambda expressions of properties to load alike l => l.BaseType</param>
        public static List GetListByTitle(this Web web, string listTitle, params Expression<Func<List, object>>[] expressions)
        {
            var baseExpressions = new List<Expression<Func<List, object>>> { l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, l => l.Hidden, l => l.RootFolder };

            if (expressions != null && expressions.Any())
            {
                baseExpressions.AddRange(expressions);
            }
            if (string.IsNullOrEmpty(listTitle))
            {
                throw (listTitle == null)
                    ? new ArgumentNullException(nameof(listTitle))
                    : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(listTitle));
            }
            var query = web.Lists.IncludeWithDefaultProperties(baseExpressions.ToArray());
            var lists = web.Context.LoadQuery(query.Where(l => l.Title == listTitle));
            web.Context.ExecuteQueryRetry();
            return lists.FirstOrDefault();
        }

        /// <summary>
        /// Get list by using Url
        /// </summary>
        /// <param name="web">Web (site) to be processed</param>
        /// <param name="webRelativeUrl">Url of list relative to the web (site), e.g. lists/testlist</param>
        /// <param name="expressions">Additional list of lambda expressions of properties to load alike l => l.BaseType</param>
        /// <returns>Returns list if found, null if no list is found.</returns>
        public static List GetListByUrl(this Web web, string webRelativeUrl, params Expression<Func<List, object>>[] expressions)
        {
            var baseExpressions = new List<Expression<Func<List, object>>> { l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, l => l.Hidden, l => l.RootFolder };

            if (expressions != null && expressions.Any())
            {
                baseExpressions.AddRange(expressions);
            }
            if (string.IsNullOrEmpty(webRelativeUrl))
                throw new ArgumentNullException(nameof(webRelativeUrl));

            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }
            var listServerRelativeUrl = UrlUtility.Combine(web.ServerRelativeUrl, webRelativeUrl);

            var foundList = web.GetList(listServerRelativeUrl);
            web.Context.Load(foundList, baseExpressions.ToArray());
            try
            {
                web.Context.ExecuteQueryRetry();
            }
            catch (ServerException se)
            {
                if (se.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    foundList = null;
                }
                else
                {
                    throw;
                }
            }

            return foundList;
        }

        /// <summary>
        /// Gets the publishing pages library of the web based on site language
        /// </summary>
        /// <param name="web">The web.</param>
        /// <returns>The publishing pages library. Returns null if library was not found.</returns>
        /// <exception cref="System.InvalidOperationException">
        /// Could not load pages library URL name from 'cmscore' resources file.
        /// </exception>
        public static List GetPagesLibrary(this Web web)
        {
            if (web == null) throw new ArgumentNullException(nameof(web));

            var context = web.Context;
            int language = (int)web.EnsureProperty(w => w.Language);

            var result = Utilities.Utility.GetLocalizedString(context, "$Resources:List_Pages_UrlName", "osrvcore", language);
            context.ExecuteQueryRetry();
            string pagesLibraryName = new Regex(@"['´`]").Replace(result.Value, "");

            if (string.IsNullOrEmpty(pagesLibraryName))
            {
                throw new InvalidOperationException("Could not load pages library URL name from 'cmscore' resources file.");
            }

            return web.GetListByUrl(pagesLibraryName) ?? web.GetListByTitle(pagesLibraryName);
        }

        /// <summary>
        /// Gets the web relative URL.
        /// Allow users to get the web relative URL of a list.  
        /// This is useful when exporting lists as it can then be used as a parameter to Web.GetListByUrl().
        /// </summary>
        /// <param name="list">The list to export the URL of.</param>
        /// <returns>The web relative URL of the list.</returns>
        public static string GetWebRelativeUrl(this List list)
        {
            list.EnsureProperties(l => l.RootFolder, l => l.ParentWebUrl);
            return GetWebRelativeUrl(list.RootFolder.ServerRelativeUrl, list.ParentWebUrl);
        }

        /// <summary>
        /// Gets the web relative URL.
        /// </summary>
        /// <param name="listRootFolderServerRelativeUrl">The list root folder server relative URL.</param>
        /// <param name="parentWebServerRelativeUrl">The parent web server relative URL.</param>
        /// <returns>The web relative URL.</returns>
        /// <exception cref="Exception">Cannot establish web relative URL from the list root folder URI and the parent web URI.</exception>
        private static string GetWebRelativeUrl(string listRootFolderServerRelativeUrl, string parentWebServerRelativeUrl)
        {
            var sanitisedListRootFolderServerRelativeUrl = listRootFolderServerRelativeUrl.Trim(UrlDelimiters);
            var sanitisedParentWebServerRelativeUrl = parentWebServerRelativeUrl.Trim(UrlDelimiters);

            if (!sanitisedListRootFolderServerRelativeUrl.StartsWith(sanitisedParentWebServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception(string.Format(CoreResources.ListExtensions_GetWebRelativeUrl, listRootFolderServerRelativeUrl, parentWebServerRelativeUrl));
            }

            var listWebRelativeUrl = sanitisedListRootFolderServerRelativeUrl.Substring(sanitisedParentWebServerRelativeUrl.Length);

            return listWebRelativeUrl.Trim(UrlDelimiters);
        }

        #region List Permissions

        /// <summary>
        /// Set custom permission to the list
        /// </summary>
        /// <param name="list">List on which permission to be set</param>
        /// <param name="user">Built in user</param>
        /// <param name="roleType">Role type</param>
        public static void SetListPermission(this List list, BuiltInIdentity user, RoleType roleType)
        {
            Principal permissionEntity = null;

            // Get the web for list
            var web = list.ParentWeb;
            list.Context.Load(web);
            list.Context.ExecuteQueryRetry();

            switch (user)
            {
                case BuiltInIdentity.Everyone:
                    {
                        permissionEntity = web.EnsureUser("c:0(.s|true");
                        break;
                    }
                case BuiltInIdentity.EveryoneButExternalUsers:
                    {
#if !NETSTANDARD2_0
                        string userIdentity = $"c:0-.f|rolemanager|spo-grid-all-users/{web.GetAuthenticationRealm()}";
                        permissionEntity = web.EnsureUser(userIdentity);
                        break;
#else
                        throw new Exception("Not Supported");
#endif
                    }
            }

            list.SetListPermission(permissionEntity, roleType);
        }

        /// <summary>
        /// Set custom permission to the list
        /// </summary>
        /// <param name="list">List on which permission to be set</param>
        /// <param name="principal">SharePoint Group or User</param>
        /// <param name="roleType">Role type</param>
        public static void SetListPermission(this List list, Principal principal, RoleType roleType)
        {
            // Get the web for list
            var web = list.ParentWeb;
            list.Context.Load(web);
            list.Context.ExecuteQueryRetry();

            // Stop inheriting permissions
            list.BreakRoleInheritance(true, false);

            // Get role type
            var roleDefinition = web.RoleDefinitions.GetByType(roleType);
            var rdbColl = new RoleDefinitionBindingCollection(web.Context) { roleDefinition };

            // Set custom permission to the list
            list.RoleAssignments.Add(principal, rdbColl);
            list.Context.ExecuteQueryRetry();
        }

        #endregion

        #region List view

        /// <summary>
        /// Creates list views based on specific xml structure from file
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="listUrl">List Url</param>
        /// <param name="filePath">Path of the file</param>
        public static void CreateViewsFromXMLFile(this Web web, string listUrl, string filePath)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException(nameof(filePath));

            var xd = new XmlDocument();
            xd.Load(filePath);
            CreateViewsFromXML(web, listUrl, xd);
        }

        /// <summary>
        /// Creates views based on specific xml structure from string
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="listUrl">List Url</param>
        /// <param name="xmlString">Path of the file</param>
        public static void CreateViewsFromXMLString(this Web web, string listUrl, string xmlString)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));

            if (string.IsNullOrEmpty(xmlString))
                throw new ArgumentNullException(nameof(xmlString));

            XmlDocument xd = new XmlDocument();
            xd.LoadXml(xmlString);
            CreateViewsFromXML(web, listUrl, xd);
        }

        /// <summary>
        /// Create list views based on xml structure loaded to memory
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="listUrl">List Url</param>
        /// <param name="xmlDoc">XmlDocument object</param>
        public static void CreateViewsFromXML(this Web web, string listUrl, XmlDocument xmlDoc)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));

            if (xmlDoc == null)
                throw new ArgumentNullException(nameof(xmlDoc));

            // Get instances to the list
            List list = web.GetList(listUrl);
            web.Context.Load(list);
            web.Context.ExecuteQueryRetry();

            // Execute the actual xml based creation
            list.CreateViewsFromXML(xmlDoc);
        }

        /// <summary>
        /// Create list views based on specific xml structure in external file
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="filePath">Path of the file</param>
        public static void CreateViewsFromXMLFile(this List list, string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException(nameof(filePath));

            if (!System.IO.File.Exists(filePath))
                throw new FileNotFoundException(filePath);

            XmlDocument xd = new XmlDocument();
            xd.Load(filePath);
            list.CreateViewsFromXML(xd);
        }

        /// <summary>
        /// Create list views based on specific xml structure in string 
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="xmlString">XML string to create view</param>
        public static void CreateViewsFromXMLString(this List list, string xmlString)
        {
            if (string.IsNullOrEmpty(xmlString))
                throw new ArgumentNullException(nameof(xmlString));

            XmlDocument xd = new XmlDocument();
            xd.LoadXml(xmlString);
            list.CreateViewsFromXML(xd);
        }

        /// <summary>
        /// Actual implementation of the view creation logic based on given xml
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="xmlDoc">XmlDocument object</param>
        public static void CreateViewsFromXML(this List list, XmlDocument xmlDoc)
        {
            if (xmlDoc == null)
                throw new ArgumentNullException(nameof(xmlDoc));

            // Convert base type to string value used in the xml structure
            var listType = list.BaseType.ToString();
            // Get only relevant list views for matching base list type
            var listViews = xmlDoc.SelectNodes("ListViews/List[@Type='" + listType + "']/View");

            foreach (XmlNode view in listViews)
            {
                string name = view.Attributes["Name"].Value;
                ViewType type = (ViewType)Enum.Parse(typeof(ViewType), view.Attributes["ViewTypeKind"].Value);
                string[] viewFields = view.Attributes["ViewFields"].Value.Split(',');
                uint rowLimit = uint.Parse(view.Attributes["RowLimit"].Value);
                bool defaultView = bool.Parse(view.Attributes["DefaultView"].Value);
                string query = view.SelectSingleNode("./ViewQuery").InnerText;

                //Create View
                list.CreateView(name, type, viewFields, rowLimit, defaultView, query);
            }
        }

        /// <summary>
        /// Create view to existing list
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="viewName">Name of the view</param>
        /// <param name="viewType">Type of the view</param>
        /// <param name="viewFields">Fields of the view</param>
        /// <param name="rowLimit">Row limit of the view</param>
        /// <param name="setAsDefault">Set as default view</param>
        /// <param name="query">Query for view creation</param>        
        /// <param name="personal">Personal View</param>
        /// <param name="paged">Paged view</param>        
        public static View CreateView(this List list,
            string viewName,
            ViewType viewType,
            string[] viewFields,
            uint rowLimit,
            bool setAsDefault,
            string query = null,
            bool personal = false,
            bool paged = false)
        {
            if (string.IsNullOrEmpty(viewName))
                throw new ArgumentNullException(nameof(viewName));

            var viewCreationInformation = new ViewCreationInformation
            {
                Title = viewName,
                ViewTypeKind = viewType,
                RowLimit = rowLimit,
                ViewFields = viewFields,
                PersonalView = personal,
                SetAsDefaultView = setAsDefault,
                Paged = paged
            };
            if (!string.IsNullOrEmpty(query))
            {
                viewCreationInformation.Query = query;
            }

            var view = list.Views.Add(viewCreationInformation);
            list.Context.Load(view);
            list.Context.ExecuteQueryRetry();

            return view;
        }

        /// <summary>
        /// Gets a view by Id
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="id">Id to the view to extract</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>returns null if not found</returns>
        public static View GetViewById(this List list, Guid id, params Expression<Func<View, object>>[] expressions)
        {

            id.ValidateNotNullOrEmpty(nameof(id));

            try
            {
                var view = list.Views.GetById(id);
                if (expressions != null && expressions.Any())
                {
                    list.Context.Load(view, expressions);
                }
                else
                {
                    list.Context.Load(view);
                }
                list.Context.ExecuteQueryRetry();

                return view;
            }
            catch (ServerException)
            {
                return null;
            }
        }

        /// <summary>
        /// Gets a view by Name
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="name">Name of the view</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>returns null if not found</returns>
        public static View GetViewByName(this List list, string name, params Expression<Func<View, object>>[] expressions)
        {
            name.ValidateNotNullOrEmpty(nameof(name));

            try
            {
                var view = list.Views.GetByTitle(name);
                if (expressions != null && expressions.Any())
                {
                    list.Context.Load(view, expressions);
                }
                else
                {
                    list.Context.Load(view);
                }
                list.Context.ExecuteQueryRetry();

                return view;
            }
            catch (ServerException)
            {
                return null;
            }

        }
        #endregion

        private static void SetDefaultColumnValuesImplementation(this List list, IEnumerable<IDefaultColumnValue> columnValues)
        {
            if (columnValues == null || !columnValues.Any()) return;
            using (var clientContext = list.Context as ClientContext)
            {
                try
                {
                    var values = columnValues.ToList<IDefaultColumnValue>();

                    clientContext.Load(list.RootFolder, r => r.ServerRelativeUrl);
                    clientContext.ExecuteQueryRetry();

                    var xMetadataDefaults = new XElement("MetadataDefaults");

                    while (values.Any())
                    {
                        // Get the first entry 
                        IDefaultColumnValue defaultColumnValue = values.First();
                        var path = defaultColumnValue.FolderRelativePath;
                        if (string.IsNullOrEmpty(path))
                        {
                            // Assume root folder
                            path = "/";
                        }
                        path = path.Equals("/") ? list.RootFolder.ServerRelativeUrl : UrlUtility.Combine(list.RootFolder.ServerRelativeUrl, path);
                        // Find all in the same path:
                        var defaultColumnValuesInSamePath = columnValues.Where(x => x.FolderRelativePath == defaultColumnValue.FolderRelativePath);
#if !NETSTANDARD2_0
                        path = Utilities.HttpUtility.UrlPathEncode(path, false);
#else
                        path = System.Web.HttpUtility.UrlEncode(path);
#endif

                        var xATag = new XElement("a", new XAttribute("href", path));

                        foreach (var defaultColumnValueInSamePath in defaultColumnValuesInSamePath)
                        {
                            var fieldName = defaultColumnValueInSamePath.FieldInternalName;
                            var fieldStringBuilder = new StringBuilder();

                            var termValue = defaultColumnValueInSamePath as DefaultColumnTermValue;
                            if (termValue != null)
                            {
                                // Term value
                                foreach (var term in termValue.Terms)
                                {
                                    string wssId = string.Empty;
                                    if (!term.IsPropertyAvailable("Id"))
                                    {
                                        term.EnsureProperties(t => t.Id, t => t.Name);
                                    }
                                    if (term.IsPropertyAvailable("CustomProperties"))
                                    {
                                        term.CustomProperties.TryGetValue("WssId", out wssId);
                                    }

                                    if (string.IsNullOrEmpty(wssId)) wssId = "-1";
                                    fieldStringBuilder.AppendFormat("{0};#{1}|{2};#", wssId, term.Name, term.Id);
                                }
                                var xDefaultValue = new XElement("DefaultValue", new XAttribute("FieldName", fieldName));
                                var fieldString = fieldStringBuilder.ToString().TrimEnd(';', '#');
                                xDefaultValue.SetValue(fieldString);
                                xATag.Add(xDefaultValue);
                            }
                            else
                            {
                                // Text value
                                var fieldString = fieldStringBuilder.Append(((DefaultColumnTextValue)defaultColumnValueInSamePath).Text);
                                var xDefaultValue = new XElement("DefaultValue", new XAttribute("FieldName", fieldName));
                                xDefaultValue.SetValue(fieldString);
                                xATag.Add(xDefaultValue);
                            }

                            values.Remove(defaultColumnValueInSamePath);
                        }
                        xMetadataDefaults.Add(xATag);
                    }

                    var formsFolder = GetFormsFolderFromList(list, clientContext);
                    if (formsFolder != null)
                    {
                        var xmlSb = new StringBuilder();
                        XmlWriterSettings xmlSettings = new XmlWriterSettings
                        {
                            OmitXmlDeclaration = true,
                            NewLineHandling = NewLineHandling.None,
                            Indent = false
                        };

                        using (var xmlWriter = XmlWriter.Create(xmlSb, xmlSettings))
                        {
                            xMetadataDefaults.Save(xmlWriter);
                        }

                        var objFileInfo = new FileCreationInformation
                        {
                            Url = "client_LocationBasedDefaults.html",
                            ContentStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlSb.ToString())),
                            Overwrite = true
                        };

                        formsFolder.Files.Add(objFileInfo);
                        clientContext.ExecuteQueryRetry();
                    }

                    // Add the event receiver if not already there
                    if (list.GetEventReceiverByName("LocationBasedMetadataDefaultsReceiver ItemAdded") == null)
                    {
                        EventReceiverDefinitionCreationInformation eventCi = new EventReceiverDefinitionCreationInformation();
                        eventCi.Synchronization = EventReceiverSynchronization.Synchronous;
                        eventCi.EventType = EventReceiverType.ItemAdded;
#if SP2013
                        eventCi.ReceiverAssembly = "Microsoft.Office.DocumentManagement, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c";
#else
                        eventCi.ReceiverAssembly = "Microsoft.Office.DocumentManagement, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c";
#endif
                        eventCi.ReceiverClass = "Microsoft.Office.DocumentManagement.LocationBasedMetadataDefaultsReceiver";
                        eventCi.ReceiverName = "LocationBasedMetadataDefaultsReceiver ItemAdded";
                        eventCi.SequenceNumber = 1000;

                        list.EventReceivers.Add(eventCi);

                        list.Update();

                        clientContext.ExecuteQueryRetry();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error applying default column values", ex);
                }
            }
        }

        /// <summary>
        /// <para>Sets default values for column values.</para>
        /// <para>In order to for instance set the default Enterprise Metadata keyword field to a term, add the enterprise metadata keyword to a library (internal name "TaxKeyword")</para>
        /// <para> </para>
        /// <para>Column values are defined by the DefaultColumnValue class that has 3 properties:</para>
        /// <para>RelativeFolderPath : / to set a default value for the root of the document library, or /foldername to specify a subfolder</para>
        /// <para>FieldInternalName : The name of the field to set. For instance "TaxKeyword" to set the Enterprise Metadata field</para>
        /// <para>Terms : A collection of Taxonomy terms to set</para>
        /// <para></para>
        /// <para>Supported column types: Metadata, Text, Choice, MultiChoice, People, Boolean, DateTime, Number, Currency</para>
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="columnValues">Column Values</param>
        public static void SetDefaultColumnValues(this List list, IEnumerable<IDefaultColumnValue> columnValues)
        {

            using (var clientContext = (ClientContext)list.Context)
            {
                clientContext.Load(list.RootFolder, r => r.ServerRelativeUrl);
                clientContext.ExecuteQueryRetry();
                // Check if default values file is present
                var formsFolder = GetFormsFolderFromList(list, clientContext);
                List<IDefaultColumnValue> existingValues = new List<IDefaultColumnValue>();

                if (formsFolder != null)
                {
                    var configFile = formsFolder.Files.GetByUrl("client_LocationBasedDefaults.html");
                    clientContext.Load(configFile, c => c.Exists);
                    bool fileExists = false;
                    try
                    {
                        clientContext.ExecuteQueryRetry();
                        fileExists = true;
                    }
                    catch { }

                    if (fileExists)
                    {
                        var streamResult = configFile.OpenBinaryStream();
                        clientContext.ExecuteQueryRetry();
                        XDocument document = XDocument.Load(streamResult.Value);
                        var values = from a in document.Descendants("a") select a;
                        Dictionary<string, Field> fieldCache = new Dictionary<string, Field>();

                        foreach (var value in values)
                        {
                            var href = value.Attribute("href").Value;
#if !NETSTANDARD2_0
                            href = Utilities.HttpUtility.UrlKeyValueDecode(href);
#else
                            href = System.Web.HttpUtility.UrlDecode(href);
#endif
                            href = href.Replace(list.RootFolder.ServerRelativeUrl, "/").Replace("//", "/");
                            var defaultValues = from d in value.Descendants("DefaultValue") select d;
                            foreach (var defaultValue in defaultValues)
                            {
                                var fieldName = defaultValue.Attribute("FieldName").Value;
                                Field field;
                                if (!fieldCache.TryGetValue(fieldName, out field))
                                {
                                    field = list.Fields.GetByInternalNameOrTitle(fieldName);
                                    clientContext.Load(field);
                                    clientContext.ExecuteQueryRetry();
                                    fieldCache.Add(fieldName, field);
                                }

                                if (field.FieldTypeKind == FieldType.Text ||
                                    field.FieldTypeKind == FieldType.Choice ||
                                    field.FieldTypeKind == FieldType.MultiChoice ||
                                    field.FieldTypeKind == FieldType.User ||
                                    field.FieldTypeKind == FieldType.Boolean ||
                                    field.FieldTypeKind == FieldType.DateTime ||
                                    field.FieldTypeKind == FieldType.Number ||
                                    field.FieldTypeKind == FieldType.Currency
                                    )
                                {
                                    var textValue = defaultValue.Value;

                                    if (field.FieldTypeKind == FieldType.User && !textValue.Contains(";#"))
                                    {
                                        Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ListExtensions_IncorrectValueFormat);
                                        continue;
                                    }

                                    var defaultColumnTextValue = new DefaultColumnTextValue()
                                    {
                                        FieldInternalName = fieldName,
                                        FolderRelativePath = href,
                                        Text = textValue
                                    };
                                    existingValues.Add(defaultColumnTextValue);
                                }
                                else
                                {
                                    var termsIdentifier = defaultValue.Value;

                                    var terms = termsIdentifier.Split(new[] { ";#" }, StringSplitOptions.None);
                                    var existingTerms = new List<Term>();
                                    for (int q = 0; q < terms.Length; q += 2)
                                    {
                                        var wssId = terms[q];
                                        var splitData = terms[q + 1].Split(new char[] { '|' });

                                        var termName = splitData[0];
                                        var termIdString = splitData[1];

                                        Term perfTerm = HydrateTermFromText(clientContext, termIdString, termName, wssId);
                                        existingTerms.Add(perfTerm);
                                    }

                                    var defaultColumnTermValue = new DefaultColumnTermValue()
                                    {
                                        FieldInternalName = fieldName,
                                        FolderRelativePath = href
                                    };
                                    existingTerms.ForEach(t => defaultColumnTermValue.Terms.Add(t));

                                    existingValues.Add(defaultColumnTermValue);
                                }
                            }
                        }
                    }
                }

                var termsList = columnValues.Union(existingValues, new DefaultColumnTermValueComparer()).ToList();
                list.SetDefaultColumnValuesImplementation(termsList);
            }
        }

        private static Term HydrateTermFromText(ClientContext clientContext, string termIdString, string termName, string wssId)
        {
            if (string.IsNullOrEmpty(wssId)) wssId = "-1";
            Term perfTerm = new Term(clientContext, null);
            var prop = perfTerm.GetType().GetProperty("ObjectData", BindingFlags.Instance | BindingFlags.NonPublic);
            ClientObjectData data = (ClientObjectData)prop.GetValue(perfTerm);
            data.Properties["Id"] = Guid.Parse(termIdString);
            data.Properties["Name"] = termName;
            data.Properties["CustomProperties"] = new Dictionary<string, string>();
            perfTerm.CustomProperties.Add("WssId", wssId);
            return perfTerm;
        }

        private static Folder GetFormsFolderFromList(List list, ClientContext clientContext)
        {
            Folder formsFolder = null;
            try
            {
                formsFolder = list.ParentWeb.GetFolderByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/Forms");
                clientContext.ExecuteQueryRetry();
            }
            catch (FileNotFoundException)
            {
                // eat the exception
                return null;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName.Equals("System.IO.FileNotFoundException"))
                {
                    // eat the exception
                    return null;
                }
                else
                {
                    throw ex;
                }
            }
            return formsFolder;
        }

        /// <summary>
        /// <para>Sets default values for column values.</para>
        /// <para>In order to for instance set the default Enterprise Metadata keyword field to a term, add the enterprise metadata keyword to a library (internal name "TaxKeyword")</para>
        /// <para> </para>
        /// <para>Column values are defined by the DefaultColumnValue class that has 3 properties:</para>
        /// <para>RelativeFolderPath : / to set a default value for the root of the document library, or /foldername to specify a subfolder</para>
        /// <para>FieldInternalName : The name of the field to set. For instance "TaxKeyword" to set the Enterprise Metadata field</para>
        /// <para>Terms : A collection of Taxonomy terms to set</para>
        /// <para></para>
        /// <para>Supported column types: Metadata, Text, Choice, MultiChoice, People, Boolean, DateTime, Number, Currency</para>
        /// </summary>
        /// <param name="list">The list to process.</param>
        /// <param name="columnValues">The default column values.</param>
        /// <param name="overwriteExistingDefaultColumnValues">If true, the currrent default column values will be overwritten.</param>
        public static void SetDefaultColumnValues(this List list, IEnumerable<IDefaultColumnValue> columnValues, bool overwriteExistingDefaultColumnValues)
        {
            if (overwriteExistingDefaultColumnValues)
            {
                list.SetDefaultColumnValuesImplementation(columnValues);
            }
            else
            {
                list.SetDefaultColumnValues(columnValues);
            }
        }

        /// <summary>
        /// Remove all default column values that are defined for this list.
        /// </summary>
        /// <param name="list">The list to process.</param>
        public static void ClearDefaultColumnValues(this List list)
        {
            var defaultValuesFileName = "client_LocationBasedDefaults.html";
            using (var clientContext = (ClientContext)list.Context)
            {
                clientContext.Load(list.RootFolder);
                clientContext.Load(list.RootFolder.Folders);
                clientContext.ExecuteQueryRetry();

                // Check if default values file is present
                var formsFolder = list.RootFolder.Folders.FirstOrDefault(x => x.Name == "Forms");
                var configFile = formsFolder.Files.GetByUrl(defaultValuesFileName);
                clientContext.Load(configFile, c => c.Exists);
                bool fileExists = false;
                try
                {
                    clientContext.ExecuteQueryRetry();
                    fileExists = true;
                }
                catch
                {
                    // Do nothing here
                }

                if (fileExists)
                {
                    configFile.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
            }
        }

        /// <summary>
        /// Removes the provided default column values from the specified folder(s) from list, if they were set.
        /// </summary>
        /// <param name="list">The list to process.</param>
        /// <param name="columnValues">The default column values that must be cleared.</param>
        public static void ClearDefaultColumnValues(this List list, IEnumerable<IDefaultColumnValue> columnValues)
        {
            var defaultValuesFileName = "client_LocationBasedDefaults.html";
            using (var clientContext = (ClientContext)list.Context)
            {
                clientContext.Load(list.RootFolder);
                clientContext.Load(list.RootFolder.Folders);
                clientContext.ExecuteQueryRetry();
                TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(clientContext);
                // Check if default values file is present
                var formsFolder = list.RootFolder.Folders.FirstOrDefault(x => x.Name == "Forms");
                List<IDefaultColumnValue> remainingValues = new List<IDefaultColumnValue>();

                if (formsFolder != null)
                {
                    var configFile = formsFolder.Files.GetByUrl(defaultValuesFileName);
                    clientContext.Load(configFile, c => c.Exists);
                    bool fileExists = false;
                    try
                    {
                        clientContext.ExecuteQueryRetry();
                        fileExists = true;
                    }
                    catch { }

                    if (fileExists)
                    {
                        var streamResult = configFile.OpenBinaryStream();
                        clientContext.ExecuteQueryRetry();
                        XDocument document = XDocument.Load(streamResult.Value);
                        var values = from a in document.Descendants("a") select a;

                        foreach (var value in values)
                        {
                            var href = value.Attribute("href").Value;
#if !NETSTANDARD2_0
                            href = Utilities.HttpUtility.UrlKeyValueDecode(href);
#else
                            href = System.Web.HttpUtility.UrlDecode(href);
#endif

                            href = href.Replace(list.RootFolder.ServerRelativeUrl, "/").Replace("//", "/");
                            var defaultValues = from d in value.Descendants("DefaultValue") select d;
                            foreach (var defaultValue in defaultValues)
                            {
                                var fieldName = defaultValue.Attribute("FieldName").Value;

                                var field = list.Fields.GetByInternalNameOrTitle(fieldName);
                                clientContext.Load(field);
                                clientContext.ExecuteQueryRetry();
                                if (field.FieldTypeKind == FieldType.Text ||
                                    field.FieldTypeKind == FieldType.Choice ||
                                    field.FieldTypeKind == FieldType.MultiChoice ||
                                    field.FieldTypeKind == FieldType.User ||
                                    field.FieldTypeKind == FieldType.Boolean ||
                                    field.FieldTypeKind == FieldType.DateTime ||
                                    field.FieldTypeKind == FieldType.Number ||
                                    field.FieldTypeKind == FieldType.Currency
                                    )
                                {
                                    var textValue = defaultValue.Value;

                                    if (field.FieldTypeKind == FieldType.User && !textValue.Contains(";#"))
                                    {
                                        Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ListExtensions_IncorrectValueFormat);
                                        continue;
                                    }

                                    var defaultColumnTextValue = new DefaultColumnTextValue()
                                    {
                                        FieldInternalName = fieldName,
                                        FolderRelativePath = href,
                                        Text = textValue
                                    };

                                    bool shouldBeKept = columnValues
                                        .FirstOrDefault(c =>
                                            c.FieldInternalName == defaultColumnTextValue.FieldInternalName && c.FolderRelativePath == defaultColumnTextValue.FolderRelativePath
                                        ) == null;
                                    if (shouldBeKept == true)
                                    {
                                        remainingValues.Add(defaultColumnTextValue);
                                    }
                                }
                                else
                                {
                                    var termsIdentifier = defaultValue.Value;

                                    var terms = termsIdentifier.Split(new string[] { ";#" }, StringSplitOptions.None);

                                    var existingTerms = new List<Term>();
                                    for (int q = 1; q < terms.Length; q++)
                                    {
                                        var termIdString = terms[q].Split(new char[] { '|' })[1];
                                        var term = taxSession.GetTerm(new Guid(termIdString));
                                        clientContext.Load(term, t => t.Id, t => t.Name);
                                        clientContext.ExecuteQueryRetry();
                                        existingTerms.Add(term);
                                        q++; // Skip one
                                    }

                                    var defaultColumnTermValue = new DefaultColumnTermValue()
                                    {
                                        FieldInternalName = fieldName,
                                        FolderRelativePath = href,
                                    };
                                    existingTerms.ForEach(t => defaultColumnTermValue.Terms.Add(t));

                                    bool shouldBeKept = columnValues
                                        .FirstOrDefault(c =>
                                            c.FieldInternalName == defaultColumnTermValue.FieldInternalName && c.FolderRelativePath == defaultColumnTermValue.FolderRelativePath
                                        ) == null;
                                    if (shouldBeKept == true)
                                    {
                                        remainingValues.Add(defaultColumnTermValue);
                                    }
                                }
                            }

                        }
                    }
                }

                list.SetDefaultColumnValuesImplementation(remainingValues);
            }
        }

        /// <summary>
        /// <para>Gets default values for column values.</para>
        /// <para></para>
        /// <para>The returned list contains one dictionary per default setting per folder.</para>
        /// <para>Each dictionary has the following keys set: Path, Field, Value</para>
        /// <para></para>
        /// <para>Path: Relative path to the library/folder</para>
        /// <para>Field: Internal name of the field which has a default value</para>
        /// <para>Value: The default value for the field</para>
        /// </summary>
        /// <param name="list">List to process</param>
        public static List<Dictionary<string, string>> GetDefaultColumnValues(this List list)
        {
            using (var clientContext = (ClientContext)list.Context)
            {
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQueryRetry();

                var formsFolder = GetFormsFolderFromList(list, clientContext);
                if (formsFolder != null)
                {
                    var configFile = formsFolder.Files.GetByUrl("client_LocationBasedDefaults.html");
                    clientContext.Load(configFile, c => c.Exists);
                    bool fileExists = false;
                    try
                    {
                        clientContext.ExecuteQueryRetry();
                        fileExists = true;
                    }
                    catch
                    {
                    }

                    if (fileExists)
                    {
                        var streamResult = configFile.OpenBinaryStream();
                        clientContext.ExecuteQueryRetry();
                        XDocument document = XDocument.Load(streamResult.Value);
                        var values = from a in document.Descendants("a") select a;

                        var currentDefaults = new List<Dictionary<string, string>>();
                        foreach (var value in values)
                        {
                            var href = value.Attribute("href").Value;
#if !NETSTANDARD2_0
                            href = Utilities.HttpUtility.UrlKeyValueDecode(href);
#else
                            href = System.Web.HttpUtility.UrlDecode(href);
#endif
                            href = href.Replace(list.RootFolder.ServerRelativeUrl, "/").Replace("//", "/");

                            var defaultValues = from d in value.Descendants("DefaultValue") select d;
                            foreach (var defaultValue in defaultValues)
                            {
                                var fieldName = defaultValue.Attribute("FieldName").Value;
                                var textValue = defaultValue.Value;
                                var fieldSetting = new Dictionary<string, string>
                                {
                                    ["Path"] = href,
                                    ["Field"] = fieldName,
                                    ["Value"] = textValue
                                };
                                currentDefaults.Add(fieldSetting);
                            }
                        }
                        return currentDefaults;
                    }
                }
                return null;
            }
        }

        private class DefaultColumnTermValueComparer : IEqualityComparer<IDefaultColumnValue>
        {
            public bool Equals(IDefaultColumnValue x, IDefaultColumnValue y)
            {
                if (ReferenceEquals(x, y)) return true;

                if (ReferenceEquals(x, null) || ReferenceEquals(y, null))
                    return false;

                return x.FieldInternalName == y.FieldInternalName && x.FolderRelativePath == y.FolderRelativePath;
            }

            public int GetHashCode(IDefaultColumnValue defaultValue)
            {
                if (ReferenceEquals(defaultValue, null)) return 0;

                var hashFolder = defaultValue.FolderRelativePath?.GetHashCode() ?? 0;

                var hashFieldInternalName = defaultValue.FieldInternalName.GetHashCode();

                return hashFolder ^ hashFieldInternalName;
            }
        }

        /// <summary>
        /// Queues a list for a full crawl the next incremental crawl
        /// </summary>
        /// <param name="list">List to process</param>
        public static void ReIndexList(this List list)
        {
            list.EnsureProperties(l => l.NoCrawl);
            if (list.NoCrawl)
            {
                Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ListExtensions_SkipNoCrawlLists);
                return;
            }

            const string reIndexKey = "vti_searchversion";
            var searchversion = 0;

            if (list.PropertyBagContainsKey(reIndexKey))
            {
                searchversion = (int)list.GetPropertyBagValueInt(reIndexKey, 0);
            }
            try
            {
                list.SetPropertyBagValue(reIndexKey, searchversion + 1);
            }
            catch (ServerUnauthorizedAccessException)
            {
                Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ListExtensions_SkipNoCrawlLists);
            }
        }
    }
}
