using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of View objects
    /// </summary>
    public partial class ViewCollection : ProvisioningTemplateCollection<View>
    {
        /// <summary>
        /// Constructor for ViewCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ViewCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
