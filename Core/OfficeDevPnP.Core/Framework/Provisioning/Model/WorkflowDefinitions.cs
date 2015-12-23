using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a collection of objects of type WorkflowDefinition
    /// </summary>
    public partial class WorkflowDefinitions : ProvisioningTemplateList<WorkflowDefinition>
    {
        public WorkflowDefinitions(ProvisioningTemplate parentTemplate):
            base(parentTemplate)
        {
        }
    }
}
