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
        /// <summary>
        /// Constructor for WebPartCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public WebPartCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
