using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of WebPart objects
    /// </summary>
    public partial class WebPartCollection : ProvisioningTemplateCollection<WebPart>
    {
        public WebPartCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
