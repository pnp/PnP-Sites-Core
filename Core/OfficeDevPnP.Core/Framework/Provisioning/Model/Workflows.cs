using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Workflows to provision
    /// </summary>
    public class Workflows
    {
        /// <summary>
        /// Defines the Workflows Definitions to provision
        /// </summary>
        public List<WorkflowDefinition> WorkflowDefinitions { get; set; }

        /// <summary>
        /// Defines the Workflows Subscriptions to provision
        /// </summary>
        public List<WorkflowSubscription> WorkflowSubscriptions { get; set; }
    }
}
