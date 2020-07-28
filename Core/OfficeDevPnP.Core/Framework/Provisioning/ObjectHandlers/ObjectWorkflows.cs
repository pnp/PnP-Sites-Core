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
using System.Xml.Linq;
using System.Xml.XPath;
using System.Collections;
using OfficeDevPnP.Core.Utilities;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWorkflows : ObjectContentHandlerBase
    {
        public override string Name
        {
            get { return "Workflows"; }
        }

        public override string InternalName => "Workflows";

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
                    web.EnsureProperties(w => w.Id, w => w.ServerRelativeUrl, w => w.Url);

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
                            from d in definitions.AsEnumerable()
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
                                XamlPath = d.Xaml.SaveXamlToFile(d.Id, creationInfo.FileConnector, lists),
                            }
                            );

                        foreach (var d in definitions.AsEnumerable())
                        {
                            if (d.RequiresInitiationForm)
                            {
                                PersistWorkflowForm(web, template, creationInfo, scope, d.InitiationUrl);
                            }
                            if (d.RequiresAssociationForm)
                            {
                                PersistWorkflowForm(web, template, creationInfo, scope, d.AssociationUrl);
                            }
                        }
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
                            from s in subscriptions.AsEnumerable()
                            select new Model.WorkflowSubscription(s.PropertyDefinitions.TokenizeWorkflowSubscriptionProperties(lists))
                            {
                                DefinitionId = s.DefinitionId,
                                Enabled = s.Enabled,
                                EventSourceId = s.EventSourceId != web.Id ? String.Format("{{listid:{0}}}", lists.First(l => l.Id == s.EventSourceId).Title) : null,
                                EventTypes = s.EventTypes.ToList(),
                                ManualStartBypassesActivationLimit = s.ManualStartBypassesActivationLimit,
                                Name = s.Name,
                                ListId = s.EventSourceId != web.Id ? String.Format("{{listid:{0}}}", lists.First(l => l.Id == s.EventSourceId).Title) : null,
#if !SP2013
                                ParentContentTypeId = s.ParentContentTypeId,
#endif
                                StatusFieldName = s.StatusFieldName,
                            }
                            );
#else
                        template.Workflows.WorkflowSubscriptions.AddRange(
                            from s in subscriptions.AsEnumerable()
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

        private void PersistWorkflowForm(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, String formUrl)
        {
            var fullUri = new Uri(UrlUtility.Combine(web.Url, formUrl));

            var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
            var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

            var formFile = new Model.File()
            {
                Folder = Tokenize(folderPath, web.Url),
                Src = formUrl,
                Overwrite = true,
            };

            // Add the file to the template
            template.Files.Add(formFile);

            // Persist file using connector
            PersistFile(web, creationInfo, scope, folderPath, fileName);
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
                        XElement xaml = XElement.Load(xamlStream);

                        int retryCount = 5;
                        int retryAttempts = 1;
                        int delay = 2000;

                        while (retryAttempts <= retryCount)
                        {
                            try
                            {
                                var workflowDefinition = deploymentService.GetDefinition(templateDefinition.Id);
                                web.Context.Load(workflowDefinition);
                                web.Context.ExecuteQueryRetry();

                                // If Definition does not exist, create it
                                // Prevent an exception if already exists
                                if (workflowDefinition.ServerObjectIsNull())
                                {
                                    // Create the WorkflowDefinition instance
                                    workflowDefinition = new Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition(web.Context)
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
                                        Xaml = parser.ParseXmlString(xaml.ToString()),
                                    };

                                    // Save the Workflow Definition
                                    var newDefinition = deploymentService.SaveDefinition(workflowDefinition);
                                    web.Context.ExecuteQueryRetry();

                                    // Let's publish the Workflow Definition, if needed
                                    if (templateDefinition.Published)
                                    {
                                        deploymentService.PublishDefinition(newDefinition.Value);
                                        web.Context.ExecuteQueryRetry();
                                    }
                                }
                                else
                                {
                                    // Otherwise update the XAML definition
                                    workflowDefinition.AssociationUrl = templateDefinition.AssociationUrl;
                                    workflowDefinition.Description = templateDefinition.Description;
                                    workflowDefinition.DisplayName = templateDefinition.DisplayName;
                                    workflowDefinition.FormField = templateDefinition.FormField;
                                    workflowDefinition.DraftVersion = templateDefinition.DraftVersion;
                                    workflowDefinition.InitiationUrl = templateDefinition.InitiationUrl;
                                    workflowDefinition.RequiresAssociationForm = templateDefinition.RequiresAssociationForm;
                                    workflowDefinition.RequiresInitiationForm = templateDefinition.RequiresInitiationForm;
                                    workflowDefinition.RestrictToScope = parser.ParseString(templateDefinition.RestrictToScope);
                                    workflowDefinition.RestrictToType = templateDefinition.RestrictToType != "Universal" ? templateDefinition.RestrictToType : null;
                                    workflowDefinition.Xaml = parser.ParseXmlString(xaml.ToString());

                                    var updatedDefinition = deploymentService.SaveDefinition(workflowDefinition);
                                    web.Context.ExecuteQueryRetry();

                                    // Let's publish the Workflow Definition, if needed
                                    if (templateDefinition.Published)
                                    {
                                        deploymentService.PublishDefinition(updatedDefinition.Value);
                                        web.Context.ExecuteQueryRetry();
                                    }
                                }

                                // If we are here, we have the workflow definition
                                // and we did not have any exception.
                                // Thus, we can exit from the loop
                                break;
                            }
                            catch (Exception ex)
                            {
                                // check exception is due to connection closed issue
                                if (ex is ServerException && ((ServerException)ex).ServerErrorCode == -2130575223 &&
                                    ((ServerException)ex).ServerErrorTypeName.Equals("Microsoft.SharePoint.SPException", StringComparison.InvariantCultureIgnoreCase) &&
                                    ((ServerException)ex).Message.Contains("A connection that was expected to be kept alive was closed by the server.")
                                    )
                                {
                                    WriteMessage($"Connection closed whilst adding Workflow Definition, trying again in {delay}ms", ProvisioningMessageType.Warning);

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
                        Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription workflowSubscription = null;

                        // Check if the subscription already exists before adding it, and 
                        // if already exists a subscription with the same name and with the same DefinitionId, 
                        // it is a duplicate and we just need to update it
                        string subscriptionName;
                        if (subscription.PropertyDefinitions.TryGetValue("SharePointWorkflowContext.Subscription.Name", out subscriptionName) &&
                            existingWorkflowSubscriptions.Any(s => s.PropertyDefinitions["SharePointWorkflowContext.Subscription.Name"] == subscriptionName && s.DefinitionId == subscription.DefinitionId))
                        {
                            // Thus, delete it before adding it again!
                            WriteMessage($"Workflow Subscription '{subscription.Name}' already exists. It will be updated.", ProvisioningMessageType.Warning);
                            workflowSubscription = existingWorkflowSubscriptions.FirstOrDefault((s => s.PropertyDefinitions["SharePointWorkflowContext.Subscription.Name"] == subscriptionName && s.DefinitionId == subscription.DefinitionId));
                        }

                        if (workflowSubscription != null)
                        {
                            // Update The existing subscription instead of delete the existing one.
                            // Only update the following properties
                            workflowSubscription.Enabled = subscription.Enabled;
                            workflowSubscription.EventTypes = subscription.EventTypes;
                            workflowSubscription.ManualStartBypassesActivationLimit = subscription.ManualStartBypassesActivationLimit;
                            workflowSubscription.StatusFieldName = subscription.StatusFieldName;
                        }
                        else
                        {
                            // Create the WorkflowDefinition instance
                            workflowSubscription =
                                new Microsoft.SharePoint.Client.WorkflowServices.WorkflowSubscription(web.Context)
                                {
                                    DefinitionId = subscription.DefinitionId,
                                    Enabled = subscription.Enabled,
                                    EventSourceId = (!String.IsNullOrEmpty(subscription.EventSourceId)) ? Guid.Parse(parser.ParseString(subscription.EventSourceId)) : web.Id,
                                    EventTypes = subscription.EventTypes,
#if !SP2013
                                    ParentContentTypeId = subscription.ParentContentTypeId,
#endif
                                    ManualStartBypassesActivationLimit =  subscription.ManualStartBypassesActivationLimit,
                                    Name =  subscription.Name,
                                    StatusFieldName = subscription.StatusFieldName,
                                };

                            foreach (var propertyDefinition in subscription.PropertyDefinitions
                               .Where(d => d.Key == "TaskListId" ||
                                           d.Key == "HistoryListId" ||
                                           d.Key == "SharePointWorkflowContext.Subscription.Id" ||
                                           d.Key == "SharePointWorkflowContext.Subscription.Name" ||
                                           d.Key == "CreatedBySPD" ||
                                           d.Key == "StatusColumnCreated")) // If set to "0" the status column will be created automatically.
                            {
                                workflowSubscription.SetProperty(propertyDefinition.Key, parser.ParseString(propertyDefinition.Value));
                            }
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

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return (template.Workflows != null &&
                (template.Workflows.WorkflowDefinitions.Count > 0 ||
                template.Workflows.WorkflowSubscriptions.Count > 0));
        }
    }

    internal static class WorkflowExtension
    {
        public static String SaveXamlToFile(this String xaml, Guid id, OfficeDevPnP.Core.Framework.Provisioning.Connectors.FileConnectorBase connector, ListCollection lists)
        {
            // Tokenize XAML to replace any ListId or ToListId attribute with the corresponding token
            XElement xamlDocument = XElement.Parse(xaml);
            string[] listIdAttributes = {"ListId", "ToListId"};
            
            var elements = (IEnumerable)xamlDocument.XPathEvaluate($"//child::*[@{listIdAttributes[0]}|@{listIdAttributes[1]}]");

            if (elements != null) // always true, consider removing 
            {
                foreach (var element in elements.Cast<XElement>())
                {
                    foreach (var listIdAttribute in listIdAttributes)
                    {
                        if (element.Attribute(listIdAttribute) != null)
                        {
                            var listId = element.Attribute(listIdAttribute).Value;
                            element.SetAttributeValue(listIdAttribute, TokenizeListIdProperty(listId, lists));
                        }
                    }
                }

                xaml = xamlDocument.ToString();
            }

            using (Stream mem = new MemoryStream())
            {
                using (StreamWriter sw = new StreamWriter(mem, Encoding.Unicode, 2048, true))
                {
                    sw.Write(xaml);
                }
                mem.Position = 0;

                String xamlFileName = $"{id.ToString()}.xaml";
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
                                result.Add(p.Key, $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}");
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
                returnValue = $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}";
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
                                result.Add(p.Key, $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}");
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

