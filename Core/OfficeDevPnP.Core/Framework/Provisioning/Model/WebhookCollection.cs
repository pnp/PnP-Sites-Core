using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a collection of objects of type Webhook
    /// </summary>
    public partial class WebhookCollection : BaseProvisioningTemplateObjectCollection<Webhook>
    {
        /// <summary>
        /// Constructor for WebhookCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public WebhookCollection(ProvisioningTemplate parentTemplate):
            base(parentTemplate)
        {
        }
    }
}
