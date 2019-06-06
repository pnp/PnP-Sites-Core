using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client.Social;
using Microsoft.SharePoint.Client.WorkflowServices;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class for workflow extension methods
    /// </summary>
    public static partial class WorkflowExtensions
    {
        #region Subscriptions
        /// <summary>
        /// Returns a workflow subscription for a site.
        /// </summary>
        /// <param name="web">The web to get workflow subscription</param>
        /// <param name="name">The name of the workflow subscription</param>
        /// <returns>Returns a WorkflowSubscription object</returns>
        public static WorkflowSubscription GetWorkflowSubscription(this Web web, string name)
        {
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscriptions = subscriptionService.EnumerateSubscriptions();
            var subscriptionQuery = from sub in subscriptions where sub.Name == name select sub;
            var subscriptionsResults = web.Context.LoadQuery(subscriptionQuery);
            web.Context.ExecuteQueryRetry();
            var subscription = subscriptionsResults.FirstOrDefault();
            return subscription;

        }

        /// <summary>
        /// Returns a workflow subscription
        /// </summary>
        /// <param name="web">The web to get workflow subscription</param>
        /// <param name="id">The id of the workflow subscription</param>
        /// <returns>Returns a WorkflowSubscription object</returns>
        public static WorkflowSubscription GetWorkflowSubscription(this Web web, Guid id)
        {
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscription = subscriptionService.GetSubscription(id);
            web.Context.Load(subscription);
            web.Context.ExecuteQueryRetry();
            return subscription;
        }

        /// <summary>
        /// Returns a workflow subscription (associations) for a list
        /// </summary>
        /// <param name="list">The target list</param>
        /// <param name="name">Name of workflow subscription to get</param>
        /// <returns>Returns a WorkflowSubscription object</returns>
        public static WorkflowSubscription GetWorkflowSubscription(this List list, string name)
        {
            var servicesManager = new WorkflowServicesManager(list.Context, list.ParentWeb);
            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscriptions = subscriptionService.EnumerateSubscriptionsByList(list.Id);
            var subscriptionQuery = from sub in subscriptions where sub.Name == name select sub;
            var subscriptionResults = list.Context.LoadQuery(subscriptionQuery);
            list.Context.ExecuteQueryRetry();
            var subscription = subscriptionResults.FirstOrDefault();
            return subscription;
        }

        /// <summary>
        /// Returns all the workflow subscriptions (associations) for the web and the lists of that web
        /// </summary>
        /// <param name="web">The target Web</param>
        /// <returns>Returns all WorkflowSubscriptions</returns>
        public static WorkflowSubscription[] GetWorkflowSubscriptions(this Web web)
        {
            // Get a reference to infrastructural services
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

            // Retrieve all the subscriptions (site and lists)
            var subscriptions = subscriptionService.EnumerateSubscriptions();
            web.Context.Load(subscriptions);
            web.Context.ExecuteQueryRetry();
            return subscriptions.ToArray();
        }

        /// <summary>
        /// Adds a workflow subscription
        /// </summary>
        /// <param name="list">The target list</param>
        /// <param name="workflowDefinitionName">The name of the workflow definition <seealso>
        ///         <cref>WorkflowExtensions.GetWorkflowDefinition</cref>
        ///     </seealso>
        /// </param>
        /// <param name="subscriptionName">The name of the workflow subscription to create</param>
        /// <param name="startManually">if True the workflow can be started manually</param>
        /// <param name="startOnCreate">if True the workflow will be started on item creation</param>
        /// <param name="startOnChange">if True the workflow will be started on item change</param>
        /// <param name="historyListName">the name of the history list. If not available it will be created</param>
        /// <param name="taskListName">the name of the task list. If not available it will be created</param>
        /// <param name="associationValues">the name-value pairs for workflow definition</param>
        /// <returns>Guid of the workflow subscription</returns>
        public static Guid AddWorkflowSubscription(this List list, string workflowDefinitionName, string subscriptionName, bool startManually, bool startOnCreate, bool startOnChange, string historyListName, string taskListName, Dictionary<string, string> associationValues = null)
        {
            var definition = list.ParentWeb.GetWorkflowDefinition(workflowDefinitionName, true);

            return AddWorkflowSubscription(list, definition, subscriptionName, startManually, startOnCreate, startOnChange, historyListName, taskListName, associationValues);
        }

        /// <summary>
        /// Adds a workflow subscription to a list
        /// </summary>
        /// <param name="list">The target list</param>
        /// <param name="workflowDefinition">The workflow definition. <seealso>
        ///         <cref>WorkflowExtensions.GetWorkflowDefinition</cref>
        ///     </seealso>
        /// </param>
        /// <param name="subscriptionName">The name of the workflow subscription to create</param>
        /// <param name="startManually">if True the workflow can be started manually</param>
        /// <param name="startOnCreate">if True the workflow will be started on item creation</param>
        /// <param name="startOnChange">if True the workflow will be started on item change</param>
        /// <param name="historyListName">the name of the history list. If not available it will be created</param>
        /// <param name="taskListName">the name of the task list. If not available it will be created</param>
        /// <param name="associationValues">the name-value pairs for workflow definition</param>
        /// <returns>Guid of the workflow subscription</returns>
        public static Guid AddWorkflowSubscription(this List list, WorkflowDefinition workflowDefinition, string subscriptionName, bool startManually, bool startOnCreate, bool startOnChange, string historyListName, string taskListName, Dictionary<string, string> associationValues = null)
        {
            // parameter validation
            subscriptionName.ValidateNotNullOrEmpty("subscriptionName");
            historyListName.ValidateNotNullOrEmpty("historyListName");
            taskListName.ValidateNotNullOrEmpty("taskListName");

            var historyList = list.ParentWeb.GetListByTitle(historyListName);
            if (historyList == null)
            {
                historyList = list.ParentWeb.CreateList(ListTemplateType.WorkflowHistory, historyListName, false);
                historyList.EnsureProperty(l => l.Id);
            }
            var taskList = list.ParentWeb.GetListByTitle(taskListName);
            if (taskList == null)
            {
                taskList = list.ParentWeb.CreateList(ListTemplateType.Tasks, taskListName, false);
                taskList.EnsureProperty(l => l.Id);
            }


            var sub = new WorkflowSubscription(list.Context);

            sub.DefinitionId = workflowDefinition.Id;
            sub.Enabled = true;
            sub.Name = subscriptionName;

            var eventTypes = new List<string>();
            if (startManually) eventTypes.Add("WorkflowStart");
            if (startOnCreate) eventTypes.Add("ItemAdded");
            if (startOnChange) eventTypes.Add("ItemUpdated");

            sub.EventTypes = eventTypes;

            sub.SetProperty("HistoryListId", historyList.Id.ToString());
            sub.SetProperty("TaskListId", taskList.Id.ToString());

            if (associationValues != null)
            {
                foreach (var key in associationValues.Keys)
                {
                    sub.SetProperty(key, associationValues[key]);
                }
            }

            var servicesManager = new WorkflowServicesManager(list.Context, list.ParentWeb);

            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

            var subscriptionResult = subscriptionService.PublishSubscriptionForList(sub, list.Id);

            list.Context.ExecuteQueryRetry();

            return subscriptionResult.Value;
        }



        /// <summary>
        /// Deletes the subscription
        /// </summary>
        /// <param name="subscription">the workflow subscription to delete</param>
        public static void Delete(this WorkflowSubscription subscription)
        {
            var clientContext = subscription.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);

            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

            subscriptionService.DeleteSubscription(subscription.Id);

            clientContext.ExecuteQueryRetry();
        }
        #endregion

        #region Definitions
        /// <summary>
        /// Returns a workflow definition for a site
        /// </summary>
        /// <param name="web">the target web</param>
        /// <param name="displayName">the workflow definition display name, which is displayed to users</param>
        /// <param name="publishedOnly">Defines whether to include only published definition, or all the definitions</param>
        /// <returns>Returns a WorkflowDefinition object</returns>
        public static WorkflowDefinition GetWorkflowDefinition(this Web web, string displayName, bool publishedOnly = true)
        {
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var deploymentService = servicesManager.GetWorkflowDeploymentService();
            var definitions = deploymentService.EnumerateDefinitions(publishedOnly);
            var definitionQuery = from def in definitions where def.DisplayName == displayName select def;
            var definitionResults = web.Context.LoadQuery(definitionQuery);
            web.Context.ExecuteQueryRetry();
            var definition = definitionResults.FirstOrDefault();
            return definition;
        }

        /// <summary>
        /// Returns a workflow definition
        /// </summary>
        /// <param name="web">the target web</param>
        /// <param name="id">the id of workflow definition</param>
        /// <returns>Returns a WorkflowDefinition object</returns>
        public static WorkflowDefinition GetWorkflowDefinition(this Web web, Guid id)
        {
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var deploymentService = servicesManager.GetWorkflowDeploymentService();

            var definition = deploymentService.GetDefinition(id);
            web.Context.Load(definition);
            web.Context.ExecuteQueryRetry();
            return definition;
        }

        /// <summary>
        /// Returns all the workflow definitions
        /// </summary>
        /// <param name="web">The target Web</param>
        /// <param name="publishedOnly">Defines whether to include only published definition, or all the definitions</param>
        /// <returns>Returns all WorkflowDefinitions</returns>
        public static WorkflowDefinition[] GetWorkflowDefinitions(this Web web, Boolean publishedOnly)
        {
            // Get a reference to infrastructural services
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var deploymentService = servicesManager.GetWorkflowDeploymentService();

            var definitions = deploymentService.EnumerateDefinitions(publishedOnly);
            web.Context.Load(definitions);
            web.Context.ExecuteQueryRetry();
            return definitions.ToArray();
        }

        /// <summary>
        /// Adds a workflow definition
        /// </summary>
        /// <param name="web">the target web</param>
        /// <param name="definition">the workflow definition to add</param>
        /// <param name="publish">specify true to publish workflow definition</param>
        /// <returns></returns>
        public static Guid AddWorkflowDefinition(this Web web, WorkflowDefinition definition, bool publish = true)
        {
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var deploymentService = servicesManager.GetWorkflowDeploymentService();

            WorkflowDefinition def = new WorkflowDefinition(web.Context);
            def.AssociationUrl = definition.AssociationUrl;
            def.Description = definition.Description;
            def.DisplayName = definition.DisplayName;
            def.DraftVersion = definition.DraftVersion;
            def.FormField = definition.FormField;
            def.Id = definition.Id != Guid.Empty ? definition.Id : Guid.NewGuid();
            foreach (var prop in definition.Properties)
            {
                def.SetProperty(prop.Key, prop.Value);
            }
            def.RequiresAssociationForm = definition.RequiresAssociationForm;
            def.RequiresInitiationForm = definition.RequiresInitiationForm;
            def.RestrictToScope = definition.RestrictToScope;
            def.RestrictToType = definition.RestrictToType;
            def.Xaml = definition.Xaml;

            var result = deploymentService.SaveDefinition(def);

            web.Context.ExecuteQueryRetry();

            if (publish)
            {
                deploymentService.PublishDefinition(result.Value);
                web.Context.ExecuteQueryRetry();
            }
            return result.Value;
        }
        /// <summary>
        /// Deletes a workflow definition
        /// </summary>
        /// <param name="definition">the workflow defition to delete</param>
        public static void Delete(this WorkflowDefinition definition)
        {
            var clientContext = definition.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);
            var deploymentService = servicesManager.GetWorkflowDeploymentService();
            deploymentService.DeleteDefinition(definition.Id);
            clientContext.ExecuteQueryRetry();
        }
        #endregion

        #region Instances
        /// <summary>
        /// Returns alls workflow instances for a site
        /// </summary>
        /// <param name="web">the target web</param>
        /// <returns>Returns a WorkflowInstanceCollection object</returns>
        public static WorkflowInstanceCollection GetWorkflowInstances(this Web web)
        {
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
            var instances = workflowInstanceService.EnumerateInstancesForSite();
            web.Context.Load(instances);
            web.Context.ExecuteQueryRetry();
            return instances;
        }

        /// <summary>
        /// Returns alls workflow instances for a list item
        /// </summary>
        /// <param name="web">the target web</param>
        /// <param name="item">the target list item to get workflow instances</param>
        /// <returns>Returns a WorkflowInstanceCollection object</returns>
        public static WorkflowInstanceCollection GetWorkflowInstances(this Web web, ListItem item)
        {
            var servicesManager = new WorkflowServicesManager(web.Context, web);
            var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
            var instances = workflowInstanceService.EnumerateInstancesForListItem(item.ParentList.Id, item.Id);
            web.Context.Load(instances);
            web.Context.ExecuteQueryRetry();
            return instances;
        }

        /// <summary>
        /// Returns all instances of a workflow for this subscription
        /// </summary>
        /// <param name="subscription">the workflow subscription to get instances</param>
        /// <returns>Returns a WorkflowInstanceCollection object</returns>
        public static WorkflowInstanceCollection GetInstances(this WorkflowSubscription subscription)
        {
            var clientContext = subscription.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);
            var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
            var instances = workflowInstanceService.Enumerate(subscription);
            clientContext.Load(instances);
            clientContext.ExecuteQueryRetry();
            return instances;
        }

        /// <summary>
        /// Cancels a workflow instance
        /// </summary>
        /// <param name="instance">the workflow instance to cancel</param>
        public static void CancelWorkFlow(this WorkflowInstance instance)
        {
            var clientContext = instance.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);
            var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
            workflowInstanceService.CancelWorkflow(instance);
            clientContext.ExecuteQueryRetry();
        }

        /// <summary>
        /// Resumes a workflow
        /// </summary>
        /// <param name="instance">the workflow instance to resume</param>
        public static void ResumeWorkflow(this WorkflowInstance instance)
        {
            var clientContext = instance.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);
            var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
            workflowInstanceService.ResumeWorkflow(instance);
            clientContext.ExecuteQueryRetry();
        }
        #endregion

        #region Messaging

        /// <summary>
        /// Publish a custom event to a target workflow instance
        /// </summary>
        /// <param name="instance">the workflow instance to publish event</param>
        /// <param name="eventName">The name of the target event</param>
        /// <param name="payload">The payload that will be sent to the event</param>
        public static void PublishCustomEvent(this WorkflowInstance instance, String eventName, String payload)
        {
            var clientContext = instance.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);
            var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
            workflowInstanceService.PublishCustomEvent(instance, eventName, payload);
            clientContext.ExecuteQueryRetry();
        }

        /// <summary>
        /// Starts a new instance of a workflow definition against the current web site
        /// </summary>
        /// <param name="web">The target web site</param>
        /// <param name="subscriptionName">The name of the workflow subscription to start</param>
        /// <param name="payload">Any input argument for the workflow instance</param>
        /// <returns>The ID of the just started workflow instance</returns>
        public static Guid StartWorkflowInstance(this Web web, String subscriptionName, IDictionary<String, Object> payload)
        {
            var clientContext = web.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);

            var workflowSubscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscriptions = workflowSubscriptionService.EnumerateSubscriptions();

            clientContext.Load(subscriptions, subs => subs.Where(sub => sub.Name == subscriptionName));
            clientContext.ExecuteQueryRetry();

            var subscription = subscriptions.FirstOrDefault();
            if (subscription != null)
            {
                return (StartWorkflowInstance(web, subscription.Id, payload));
            }
            else
            {
                return (Guid.Empty);
            }
        }

        /// <summary>
        /// Starts a new instance of a workflow definition against the current web site
        /// </summary>
        /// <param name="web">The target web site</param>
        /// <param name="subscriptionId">The ID of the workflow subscription to start</param>
        /// <param name="payload">Any input argument for the workflow instance</param>
        /// <returns>The ID of the just started workflow instance</returns>
        public static Guid StartWorkflowInstance(this Web web, Guid subscriptionId, IDictionary<String, Object> payload)
        {
            Guid result = Guid.Empty;

            var clientContext = web.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);

            var workflowSubscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscriptions = workflowSubscriptionService.EnumerateSubscriptions();

            clientContext.Load(subscriptions, subs => subs.Where(sub => sub.Id == subscriptionId));
            clientContext.ExecuteQueryRetry();

            var subscription = subscriptions.FirstOrDefault();
            if (subscription != null)
            {
                var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
                var startAction = workflowInstanceService.StartWorkflow(subscription, payload);
                clientContext.ExecuteQueryRetry();

                result = startAction.Value;
            }

            return (result);
        }

        /// <summary>
        /// Starts a new instance of a workflow definition against the current item
        /// </summary>
        /// <param name="item">The target item</param>
        /// <param name="subscriptionName">The name of the workflow subscription to start</param>
        /// <param name="payload">Any input argument for the workflow instance</param>
        /// <returns>The ID of the just started workflow instance</returns>
        public static Guid StartWorkflowInstance(this ListItem item, String subscriptionName, IDictionary<String, Object> payload)
        {
            var parentList = item.EnsureProperty(i => i.ParentList);

            var clientContext = item.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);

            var workflowSubscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscriptions = workflowSubscriptionService.EnumerateSubscriptionsByList(parentList.Id);

            clientContext.Load(subscriptions, subs => subs.Where(sub => sub.Name == subscriptionName));
            clientContext.ExecuteQueryRetry();

            var subscription = subscriptions.FirstOrDefault();
            if (subscription != null)
            {
                return (StartWorkflowInstance(item, subscription.Id, payload));
            }
            else
            {
                return (Guid.Empty);
            }
        }

        /// <summary>
        /// Starts a new instance of a workflow definition against the current item
        /// </summary>
        /// <param name="item">The target item</param>
        /// <param name="subscriptionId">The ID of the workflow subscription to start</param>
        /// <param name="payload">Any input argument for the workflow instance</param>
        /// <returns>The ID of the just started workflow instance</returns>
        public static Guid StartWorkflowInstance(this ListItem item, Guid subscriptionId, IDictionary<String, Object> payload)
        {
            Guid result = Guid.Empty;

            var parentList = item.EnsureProperty(i => i.ParentList);

            var clientContext = item.Context as ClientContext;
            var servicesManager = new WorkflowServicesManager(clientContext, clientContext.Web);

            var workflowSubscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscriptions = workflowSubscriptionService.EnumerateSubscriptionsByList(parentList.Id);

            clientContext.Load(subscriptions, subs => subs.Where(sub => sub.Id == subscriptionId));
            clientContext.ExecuteQueryRetry();

            var subscription = subscriptions.FirstOrDefault();
            if (subscription != null)
            {
                var workflowInstanceService = servicesManager.GetWorkflowInstanceService();
                var startAction = workflowInstanceService.StartWorkflowOnListItem(subscription, item.Id, payload);
                clientContext.ExecuteQueryRetry();

                result = startAction.Value;
            }

            return (result);
        }

        #endregion
    }
}
