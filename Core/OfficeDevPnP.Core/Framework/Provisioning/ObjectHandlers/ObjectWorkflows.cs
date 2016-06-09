using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using Microsoft.SharePoint.Client.WorkflowServices;
using System.IO;
using System.Threading;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWorkflows : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Workflows"; }
        }
        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (template.Workflows == null)
            {
                template.Workflows = new Workflows();
            }

            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (creationInfo.FileConnector == null)
                {
                    scope.LogWarning("Cannot export Workflow definitions without a FileConnector.");
                }
                else
                {
                    // Pre-load useful properties
                    web.EnsureProperty(w => w.Id);

                    // Retrieve all the lists and libraries
                    var lists = web.Lists;
                    web.Context.Load(lists);
                    web.Context.ExecuteQueryRetry();

                    // Retrieve the workflow definitions (including unpublished ones)
                    Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition[] definitions = null;

                    try
                    {
                        definitions = web.GetWorkflowDefinitions(false);
                    }
                    catch (ServerException)
                    {
                        // If there is no workflow service present in the farm this method will throw an error. 
                        // Swallow the exception
                    }

                    if (definitions != null)
                    {
                        template.Workflows.WorkflowDefinitions.AddRange(
                            from d in definitions
                            select new Model.WorkflowDefinition(d.Properties.TokenizeWorkflowDefinitionProperties(lists))
                            {
                                AssociationUrl = d.AssociationUrl,
                                Description = d.Description,
                                DisplayName = d.DisplayName,
                                DraftVersion = d.DraftVersion,
                                FormField = d.FormField,
                                Id = d.Id,
                                InitiationUrl = d.InitiationUrl,
                                Published = d.Published,
                                RequiresAssociationForm = d.RequiresAssociationForm,
                                RequiresInitiationForm = d.RequiresInitiationForm,
                                RestrictToScope = (!String.IsNullOrEmpty(d.RestrictToScope) && Guid.Parse(d.RestrictToScope) != web.Id) ? WorkflowExtension.TokenizeListIdProperty(d.RestrictToScope, lists) : null,
                                RestrictToType = !String.IsNullOrEmpty(d.RestrictToType) ? d.RestrictToType : "Universal",
                                XamlPath = d.Xaml.SaveXamlToFile(d.Id, creationInfo.FileConnector),
                            }
                            );
                    }

                    // Retrieve the workflow subscriptions
                    Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription[] subscriptions = null;

                    try
                    {
                        subscriptions = web.GetWorkflowSubscriptions();
                    }
                    catch (ServerException)
                    {
                        // If there is no workflow service present in the farm this method will throw an error. 
                        // Swallow the exception
                    }

                    if (subscriptions != null)
                    {
#if ONPREMISES
                        template.Workflows.WorkflowSubscriptions.AddRange(
                            from s in subscriptions
                            select new Model.WorkflowSubscription(s.PropertyDefinitions.TokenizeWorkflowSubscriptionProperties(lists))
                            {
                                DefinitionId = s.DefinitionId,
                                Enabled = s.Enabled,
                                EventSourceId = s.EventSourceId != web.Id ? String.Format("{{listid:{0}}}", lists.First(l => l.Id == s.EventSourceId).Title) : null,
                                EventTypes = s.EventTypes.ToList(),
                                ManualStartBypassesActivationLimit = s.ManualStartBypassesActivationLimit,
                                Name = s.Name,
                                ListId = s.EventSourceId != web.Id ? String.Format("{{listid:{0}}}", lists.First(l => l.Id == s.EventSourceId).Title) : null,
                                StatusFieldName = s.StatusFieldName,
                            }
                            );
#else
                        template.Workflows.WorkflowSubscriptions.AddRange(
                            from s in subscriptions
                            select new Model.WorkflowSubscription(s.PropertyDefinitions.TokenizeWorkflowSubscriptionProperties(lists))
                            {
                                DefinitionId = s.DefinitionId,
                                Enabled = s.Enabled,
                                EventSourceId = s.EventSourceId != web.Id ? WorkflowExtension.TokenizeListIdProperty(s.EventSourceId.ToString(), lists) : null,
                                EventTypes = s.EventTypes.ToList(),
                                ManualStartBypassesActivationLimit = s.ManualStartBypassesActivationLimit,
                                Name = s.Name,
                                ListId = s.EventSourceId != web.Id ? WorkflowExtension.TokenizeListIdProperty(s.EventSourceId.ToString(), lists) : null,
                                ParentContentTypeId = s.ParentContentTypeId,
                                StatusFieldName = s.StatusFieldName,
                            }
                            );
#endif
                    }
                }
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Get a reference to infrastructural services
                WorkflowServicesManager servicesManager = null;

                try
                {
                    servicesManager = new WorkflowServicesManager(web.Context, web);
                }
                catch (ServerException)
                {
                    // If there is no workflow service present in the farm this method will throw an error. 
                    // Swallow the exception
                }

                if (servicesManager != null)
                {
                    var deploymentService = servicesManager.GetWorkflowDeploymentService();
                    var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

                    // Pre-load useful properties
                    web.EnsureProperty(w => w.Id);

                    // Provision Workflow Definitions
                    foreach (var templateDefinition in template.Workflows.WorkflowDefinitions)
                    {
                        // Load the Workflow Definition XAML
                        Stream xamlStream = template.Connector.GetFileStream(templateDefinition.XamlPath);
                        System.Xml.Linq.XElement xaml = System.Xml.Linq.XElement.Load(xamlStream);

                        int retryCount = 5;
                        int retryAttempts = 1;
                        int delay = 2000;

                        while (retryAttempts <= retryCount)
                        {
                            try
                            {

                                // Create the WorkflowDefinition instance
                                Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition workflowDefinition =
                            new Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition(web.Context)
                            {
                                AssociationUrl = templateDefinition.AssociationUrl,
                                Description = templateDefinition.Description,
                                DisplayName = templateDefinition.DisplayName,
                                FormField = templateDefinition.FormField,
                                DraftVersion = templateDefinition.DraftVersion,
                                Id = templateDefinition.Id,
                                InitiationUrl = templateDefinition.InitiationUrl,
                                RequiresAssociationForm = templateDefinition.RequiresAssociationForm,
                                RequiresInitiationForm = templateDefinition.RequiresInitiationForm,
                                RestrictToScope = parser.ParseString(templateDefinition.RestrictToScope),
                                RestrictToType = templateDefinition.RestrictToType != "Universal" ? templateDefinition.RestrictToType : null,
                                Xaml = xaml.ToString(),
                            };

                                //foreach (var p in definition.Properties)
                                //{
                                //    workflowDefinition.SetProperty(p.Key, parser.ParseString(p.Value));
                                //}

                                // Save the Workflow Definition
                                var newDefinition = deploymentService.SaveDefinition(workflowDefinition);
                                //web.Context.Load(workflowDefinition); //not needed
                                web.Context.ExecuteQueryRetry();

                                // Let's publish the Workflow Definition, if needed
                                if (templateDefinition.Published)
                                {
                                    deploymentService.PublishDefinition(newDefinition.Value);
                                    web.Context.ExecuteQueryRetry();
                                }

                                break; // no errors so exit loop
                            }
                            catch (Exception ex)
                            {
                                // check exception is due to connection closed issue
                                if (ex is ServerException && ((ServerException)ex).ServerErrorCode == -2130575223 &&
                                    ((ServerException)ex).ServerErrorTypeName.Equals("Microsoft.SharePoint.SPException", StringComparison.InvariantCultureIgnoreCase) &&
                                    ((ServerException)ex).Message.Contains("A connection that was expected to be kept alive was closed by the server.")
                                    )
                                {
                                    WriteWarning(String.Format("Connection closed whilst adding Workflow Definition, trying again in {0}ms", delay), ProvisioningMessageType.Warning);

                                    Thread.Sleep(delay);

                                    retryAttempts++;
                                    delay = delay * 2; // double delay for next retry
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }


                    // get existing subscriptions
                    var existingWorkflowSubscriptions = web.GetWorkflowSubscriptions();

                    foreach (var subscription in template.Workflows.WorkflowSubscriptions)
                    {
                        // Check if the subscription already exists before adding it, and 
                        // if already exists a subscription with the same name and with the same DefinitionId, 
                        // it is a duplicate
                        string subscriptionName;
                        if (subscription.PropertyDefinitions.TryGetValue("SharePointWorkflowContext.Subscription.Name", out subscriptionName) &&
                            existingWorkflowSubscriptions.Any(s => s.PropertyDefinitions["SharePointWorkflowContext.Subscription.Name"] == subscriptionName && s.DefinitionId == subscription.DefinitionId))
                        {
                            // Thus, skip it!
                            WriteWarning(string.Format("Workflow Subscription '{0}' already exists. Skipping...", subscription.Name), ProvisioningMessageType.Warning);
                            continue;
                        }
#if ONPREMISES
                    // Create the WorkflowDefinition instance
                    Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription workflowSubscription =
                        new Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription(web.Context)
                        {
                            DefinitionId = subscription.DefinitionId,
                            Enabled = subscription.Enabled,
                            EventSourceId = (!String.IsNullOrEmpty(subscription.EventSourceId)) ? Guid.Parse(parser.ParseString(subscription.EventSourceId)) : web.Id,
                            EventTypes = subscription.EventTypes,
                            ManualStartBypassesActivationLimit =  subscription.ManualStartBypassesActivationLimit,
                            Name =  subscription.Name,
                            StatusFieldName = subscription.StatusFieldName,
                        };
#else
                        // Create the WorkflowDefinition instance
                        Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription workflowSubscription =
                            new Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription(web.Context)
                            {
                                DefinitionId = subscription.DefinitionId,
                                Enabled = subscription.Enabled,
                                EventSourceId = (!String.IsNullOrEmpty(subscription.EventSourceId)) ? Guid.Parse(parser.ParseString(subscription.EventSourceId)) : web.Id,
                                EventTypes = subscription.EventTypes,
                                ManualStartBypassesActivationLimit = subscription.ManualStartBypassesActivationLimit,
                                Name = subscription.Name,
                                ParentContentTypeId = subscription.ParentContentTypeId,
                                StatusFieldName = subscription.StatusFieldName,
                            };
#endif
                        foreach (var propertyDefinition in subscription.PropertyDefinitions
                            .Where(d => d.Key == "TaskListId" ||
                                        d.Key == "HistoryListId" ||
                                        d.Key == "SharePointWorkflowContext.Subscription.Id" ||
                                        d.Key == "SharePointWorkflowContext.Subscription.Name" ||
                                        d.Key == "CreatedBySPD"))
                        {
                            workflowSubscription.SetProperty(propertyDefinition.Key, parser.ParseString(propertyDefinition.Value));
                        }
                        if (!String.IsNullOrEmpty(subscription.ListId))
                        {
                            // It is a List Workflow
                            Guid targetListId = Guid.Parse(parser.ParseString(subscription.ListId));
                            subscriptionService.PublishSubscriptionForList(workflowSubscription, targetListId);
                        }
                        else
                        {
                            // It is a Site Workflow
                            subscriptionService.PublishSubscription(workflowSubscription);
                        }
                        web.Context.ExecuteQueryRetry();
                    }
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return (template.Workflows != null &&
                (template.Workflows.WorkflowDefinitions.Count > 0 ||
                template.Workflows.WorkflowSubscriptions.Count > 0));
        }
    }

    internal static class WorkflowExtension
    {
        public static String SaveXamlToFile(this String xaml, Guid id, OfficeDevPnP.Core.Framework.Provisioning.Connectors.FileConnectorBase connector)
        {
            using (Stream mem = new MemoryStream())
            {
                using (StreamWriter sw = new StreamWriter(mem, Encoding.Unicode, 2048, true))
                {
                    sw.Write(xaml);
                }
                mem.Position = 0;

                String xamlFileName = String.Format("{0}.xaml", id.ToString());
                connector.SaveFileStream(xamlFileName, mem);
                return (xamlFileName);
            }
        }

        public static Dictionary<String, String> TokenizeWorkflowDefinitionProperties(this IDictionary<String, String> properties, ListCollection lists)
        {
            Dictionary<String, String> result = new Dictionary<String, String>();
            foreach (var p in properties)
            {
                switch (p.Key)
                {
                    case "RestrictToScope":
                    case "HistoryListId":
                    case "TaskListId":
                        if (!String.IsNullOrEmpty(p.Value))
                        {
                            var list = lists.FirstOrDefault(l => l.Id == Guid.Parse(p.Value));
                            if (list != null)
                            {
                                result.Add(p.Key, String.Format("{{listid:{0}}}", list.Title));
                            }
                        }
                        break;
                    //case "SubscriptionId":
                    //case "ServerUrl":
                    //case "EncodedAbsUrl":
                    //case "MetaInfo":
                    default:
                        result.Add(p.Key, p.Value);
                        break;
                }
            }
            return (result);
        }

        public static string TokenizeListIdProperty(string listId, ListCollection lists)
        {
            var returnValue = listId;
            var list = lists.FirstOrDefault(l => l.Id == Guid.Parse(listId));
            if (list != null)
            {
                returnValue = String.Format("{{listid:{0}}}", list.Title);
            }

            return returnValue;
        }

        public static Dictionary<String, String> TokenizeWorkflowSubscriptionProperties(this IDictionary<String, String> properties, ListCollection lists)
        {
            Dictionary<String, String> result = new Dictionary<String, String>();
            foreach (var p in properties)
            {
                switch (p.Key)
                {
                    case "TaskListId":
                    case "HistoryListId":
                        if (!String.IsNullOrEmpty(p.Value))
                        {
                            var list = lists.FirstOrDefault(l => l.Id == Guid.Parse(p.Value));
                            if (list != null)
                            {
                                result.Add(p.Key, String.Format("{{listid:{0}}}", list.Title));
                            }
                        }
                        break;
                    //case "Microsoft.SharePoint.ActivationProperties.ListId":
                    //case "SharePointWorkflowContext.Subscription.Id":
                    //case "CurrentWebUri":
                    //case "SharePointWorkflowContext.Subscription.EventSourceId":
                    //case "SharePointWorkflowContext.Subscription.EventType":
                    //case "SharePointWorkflowContext.ActivationProperties.SiteId":
                    //case "SharePointWorkflowContext.ActivationProperties.WebId":
                    //case "ScopeId":
                    default:
                        result.Add(p.Key, p.Value);
                        break;
                }
            }
            return (result);
        }
    }
}

