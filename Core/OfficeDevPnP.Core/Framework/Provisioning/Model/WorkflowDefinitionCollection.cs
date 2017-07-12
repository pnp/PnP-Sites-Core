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
    public partial class WorkflowDefinitionCollection : ProvisioningTemplateCollection<WorkflowDefinition>
    {
        /// <summary>
        /// Constructor for WorkflowDefinitionCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public WorkflowDefinitionCollection(ProvisioningTemplate parentTemplate):
            base(parentTemplate)
        {
        }
    }
}
