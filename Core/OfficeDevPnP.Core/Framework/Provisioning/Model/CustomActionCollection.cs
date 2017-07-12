using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of CustomAction objects
    /// </summary>
    public partial class CustomActionCollection : ProvisioningTemplateCollection<CustomAction>
    {
        /// <summary>
        /// Constructor for CustomActionCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public CustomActionCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
