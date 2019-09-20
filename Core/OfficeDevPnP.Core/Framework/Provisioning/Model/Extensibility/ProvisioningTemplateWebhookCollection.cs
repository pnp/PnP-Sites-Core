using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// The Webhooks for the Provisioning Template
    /// </summary>
    public partial class ProvisioningTemplateWebhookCollection : BaseProvisioningTemplateObjectCollection<ProvisioningTemplateWebhook>
    {
        public ProvisioningTemplateWebhookCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
