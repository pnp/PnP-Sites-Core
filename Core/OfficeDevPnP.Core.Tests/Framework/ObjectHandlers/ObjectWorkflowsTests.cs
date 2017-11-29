using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Linq;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectWorkflowsTests
    {
        private Guid _listId; // For easy reference

        private static readonly string SampleWorkflowPath = "../../Resources/workflow.xaml";
        private static readonly string WFStatusFieldName01 = "PnP_Test_WF_Status_01";
        private static readonly string WFStatusFieldName02 = "PnP_Test_WF_Status_02";

        #region Test initialize and cleanup

        [TestInitialize]
        public void Initialize()
        {
            Console.WriteLine("ObjectWorkflowsTests.Initialise");

            using (var cc = TestCommon.CreateClientContext())
            {
                var listCI = new ListCreationInformation();
                listCI.TemplateType = (int)ListTemplateType.GenericList;
                listCI.Title = "Test_List_Workflows_" + DateTime.Now.ToFileTime();
                var list = cc.Web.Lists.Add(listCI);
                cc.Load(list);
                cc.ExecuteQueryRetry();
                _listId = list.Id;

                list.Fields.AddFieldAsXml($@"<Field ID=""F1A1715E-6C52-40DE-8403-E9AAFD0470D0"" Type=""Text"" Name=""{WFStatusFieldName01}"" 
                DisplayName=""WF Status"" Group=""PnP"" />", true, AddFieldOptions.AddFieldInternalNameHint);
                list.Fields.AddFieldAsXml($@"<Field ID=""F1A1715E-6C52-40DE-8403-E9AAFD0470D1"" Type=""Text"" Name=""{WFStatusFieldName02}"" 
                DisplayName=""WF Status"" Group=""PnP"" />", true, AddFieldOptions.AddFieldInternalNameHint);
                cc.ExecuteQueryRetry();
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            Console.WriteLine("ObjectWorkflowsTests.Cleanup");

            using (var cc = TestCommon.CreateClientContext())
            {
                // Clean up list
                var list = cc.Web.Lists.GetById(_listId);
                list.DeleteObject();
                cc.ExecuteQueryRetry();
            }
        }

        #endregion Test initialize and cleanup

        [TestMethod]
        [Timeout(5 * 60 * 1000)]
        public void UpdatWorkflowSubscription()
        {
            var template = new ProvisioningTemplate
            {
                Workflows = new Workflows(),
            };


            var wfDefinition = new WorkflowDefinition
            {
                DisplayName = "PnP Test Workflow",
                Description = "PnP Test Workflow Description",
                Id = Guid.Parse("{19100c31-d561-42c3-88e0-5214d5c584c4}"),
                Published = true,
                RestrictToType = "List",
                RestrictToScope = _listId.ToString(),
                XamlPath = SampleWorkflowPath,
                DraftVersion = "1",               
            };

            template.Workflows.WorkflowDefinitions.Add(wfDefinition);

            var wfSubscription = new WorkflowSubscription
            {
                Name = "PnP Test Workflow",
                DefinitionId = wfDefinition.Id,
                ListId = _listId.ToString(),
                Enabled = true,
                EventSourceId = _listId.ToString(),
                EventTypes = new List<string>((new string[] { "WorkflowStart" })),
                ManualStartBypassesActivationLimit = false,
                StatusFieldName = WFStatusFieldName01,
                ParentContentTypeId = "0x01"
            };

            wfSubscription.PropertyDefinitions.Add("SharePointWorkflowContext.Subscription.Id", "d21cf99d-ada1-486b-bfcf-7d58b8a56974");
            wfSubscription.PropertyDefinitions.Add("SharePointWorkflowContext.Subscription.Name", "PnPTestWorkflow_v1_0_0_WorkflowAssociation");
 
            template.Workflows.WorkflowSubscriptions.Add(wfSubscription);

            using (var ctx = TestCommon.CreateClientContext())
            {
                TokenParser parser = new TokenParser(ctx.Web, template);
                new ObjectWorkflows().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                // Update Properties
                wfSubscription.EventTypes = new List<string>((new string[] { "WorkflowStart", "ItemUpdated" }));
                wfSubscription.Enabled = !wfSubscription.Enabled;
                wfSubscription.ManualStartBypassesActivationLimit = !wfSubscription.ManualStartBypassesActivationLimit;
                wfSubscription.StatusFieldName = WFStatusFieldName02;

                // Provision Updated Workflow
                new ObjectWorkflows().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                // Check if Values of the subscription are updated
                var subscription = ctx.Web.GetWorkflowSubscription(Guid.Parse(wfSubscription.PropertyDefinitions["SharePointWorkflowContext.Subscription.Id"]));
                Assert.AreEqual(subscription.StatusFieldName, wfSubscription.StatusFieldName);
                Assert.AreEqual(subscription.Enabled, wfSubscription.Enabled);
                Assert.AreEqual(subscription.ManualStartBypassesActivationLimit, wfSubscription.ManualStartBypassesActivationLimit);
                Assert.AreEqual(subscription.EventTypes[0], wfSubscription.EventTypes[0]);
                Assert.AreEqual(subscription.EventTypes[1], wfSubscription.EventTypes[1]);
            }
        }
    }
}
