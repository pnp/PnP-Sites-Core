using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of AvailableWebTemplate objects
    /// </summary>
    public partial class AvailableWebTemplateCollection : ProvisioningTemplateCollection<AvailableWebTemplate>
    {
        public AvailableWebTemplateCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
