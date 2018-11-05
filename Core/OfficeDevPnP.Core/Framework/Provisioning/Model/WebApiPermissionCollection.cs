using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of WebApiPermission objects
    /// </summary>
    public partial class WebApiPermissionCollection : BaseProvisioningTemplateObjectCollection<WebApiPermission>
    {
        /// <summary>
        /// Constructor for WebApiPermissionCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public WebApiPermissionCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
