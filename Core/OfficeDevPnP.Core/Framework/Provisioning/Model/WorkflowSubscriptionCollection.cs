using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a collection of objects of type WorkflowSubscription
    /// </summary>
    public partial class WorkflowSubscriptionCollection : ProvisioningTemplateCollection<WorkflowSubscription>
    {
        public WorkflowSubscriptionCollection(ProvisioningTemplate parentTemplate):
            base(parentTemplate)
        {

        }
    }
}
