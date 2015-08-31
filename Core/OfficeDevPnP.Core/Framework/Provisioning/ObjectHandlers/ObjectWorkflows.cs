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
                    scope.LogWarning("Cannot export Workflow definitios without a FileConnector.");
                }
                else
                {
                    // Retrieve all the lists and libraries
                    var lists = web.Lists;
                    web.Context.Load(lists);
                    web.Context.ExecuteQuery();

                    // Retrieve the workflow definitions (including unpublished ones)
                    var definitions = web.GetWorkflowDefinitions(false);

                    template.Workflows.WorkflowDefinitions.AddRange(
                        from d in definitions
                        select new Model.WorkflowDefinition(d.Properties.ToDictionary(i => i.Key, i => i.Value))
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
                            RestrictToScope = (!String.IsNullOrEmpty(d.RestrictToScope) && Guid.Parse(d.RestrictToScope) != web.Id) ? String.Format("{{listid:{0}}}", lists.First(l => l.Id == Guid.Parse(d.RestrictToScope)).Title) : null,
                            RestrictToType = !String.IsNullOrEmpty(d.RestrictToType) ? d.RestrictToType : "Universal",
                            XamlPath = d.Xaml.SaveXamlToFile(d.Id, creationInfo.FileConnector),
                        }
                        );

                    // Retrieve the workflow subscriptions
                    var subscriptions = web.GetWorkflowSubscriptions();

                    template.Workflows.WorkflowSubscriptions.AddRange(
                        from s in subscriptions
                        select new Model.WorkflowSubscription(s.PropertyDefinitions.ToDictionary(i => i.Key, i => i.Value))
                        {
                            DefinitionId = s.DefinitionId,
                            Enabled = s.Enabled,
                            EventSourceId = s.EventSourceId != web.Id ? String.Format("{{listid:{0}}}", lists.First(l => l.Id == s.EventSourceId).Title) : null,
                            EventTypes = s.EventTypes.ToList(),
                            ManualStartBypassesActivationLimit = s.ManualStartBypassesActivationLimit,
                            Name = s.Name,
                            ListId  = s.EventSourceId != web.Id ? String.Format("{{listid:{0}}}", lists.First(l => l.Id == s.EventSourceId).Title) : null,
                            ParentContentTypeId = s.ParentContentTypeId,
                            StatusFieldName = s.StatusFieldName,
                        }
                        );
                }
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Get a reference to infrastructural services
                var servicesManager = new WorkflowServicesManager(web.Context, web);
                var deploymentService = servicesManager.GetWorkflowDeploymentService();
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

                // Provision Workflow Definitions
                foreach (var definition in template.Workflows.WorkflowDefinitions)
                {
                    // Load the Workflow Definition XAML
                    Stream xamlStream = template.Connector.GetFileStream(definition.XamlPath);
                    System.Xml.Linq.XElement xaml = System.Xml.Linq.XElement.Load(xamlStream);

                    // Create the WorkflowDefinition instance
                    Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition workflowDefinition =
                        new Microsoft.SharePoint.Client.WorkflowServices.WorkflowDefinition(web.Context)
                        {
                            AssociationUrl = definition.AssociationUrl,
                            Description = definition.Description,
                            DisplayName = definition.DisplayName,
                            FormField = definition.FormField,
                            DraftVersion = definition.DraftVersion,
                            Id = definition.Id,
                            InitiationUrl = definition.InitiationUrl,
                            RequiresAssociationForm = definition.RequiresAssociationForm,
                            RequiresInitiationForm = definition.RequiresInitiationForm,
                            RestrictToScope = parser.ParseString(definition.RestrictToScope),
                            RestrictToType = definition.RestrictToType != "Universal" ? definition.RestrictToType : null,
                            Xaml = xaml.ToString(),
                        };

                    //foreach (var p in definition.Properties)
                    //{
                    //    workflowDefinition.SetProperty(p.Key, parser.ParseString(p.Value));
                    //}

                    // Save the Workflow Definition
                    var definitionId = deploymentService.SaveDefinition(workflowDefinition);
                    web.Context.Load(workflowDefinition);
                    web.Context.ExecuteQueryRetry();

                    // Let's publish the Workflow Definition, if needed
                    if (definition.Published)
                    {
                        deploymentService.PublishDefinition(definitionId.Value);
                    }
                }

                foreach (var subscription in template.Workflows.WorkflowSubscriptions)
                {
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
                            ParentContentTypeId = subscription.ParentContentTypeId,
                            StatusFieldName = subscription.StatusFieldName,
                        };

                    //foreach (var p in subscription.PropertyDefinitions)
                    //{
                    //    workflowSubscription.SetProperty(p.Key, parser.ParseString(p.Value));
                    //}

                    if (!String.IsNullOrEmpty(subscription.ListId))
                    {
                        Guid targetListId = Guid.Parse(parser.ParseString(subscription.ListId));
                        // It is a List Workflow
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

        public static Dictionary<String, String> TokenizeProperties(this IDictionary<String, String> properties)
        {
            Dictionary<String, String> result = new Dictionary<String, String>();
            foreach (var p in properties)
            {
                switch (p.Key)
                {
                    case "RestrictToScope":
                        break;
                    case "HistoryListId":
                        break;
                    case "TaskListId":
                        break;
                    case "SubscriptionId":
                        break;
                }
            }
            return (result);
        }
    }
}

