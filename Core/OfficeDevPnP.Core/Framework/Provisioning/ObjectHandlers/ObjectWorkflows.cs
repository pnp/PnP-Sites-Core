using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using Microsoft.SharePoint.Client.WorkflowServices;

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
                //bool include = false;

                //// Retrieve the workflow definitions
                //var definitions = web.GetWorkflowDefinitions(false);
                //include = definitions.Length > 0;

                //template.Workflows.WorkflowDefinitions.AddRange(
                //    from d in definitions
                //    select new Model.WorkflowDefinition
                //    {
                //        AssociationUrl = d.AssociationUrl,
                //        Description = d.Description,
                //        DisplayName = d.DisplayName,
                //        FormField = d.FormField,
                //        Id = d.Id,
                //        InitiationUrl = d.InitiationUrl,
                //        RequiresAssociationForm = d.RequiresAssociationForm,
                //        RequiresInitiationForm = d.RequiresInitiationForm,
                //        RestrictToScope = d.RestrictToScope,
                //        RestrictToType = d.RestrictToType,
                //        XamlPath = d.Xaml,
                //    }
                //    );
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {

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
                template.Workflows.WorkflowDefinitions.Count > 0 ||
                template.Workflows.WorkflowSubscriptions.Count > 0);
        }
    }
}

