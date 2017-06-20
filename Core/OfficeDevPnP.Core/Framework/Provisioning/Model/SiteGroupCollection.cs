using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of SiteGroup objects
    /// </summary>
    public partial class SiteGroupCollection : ProvisioningTemplateCollection<SiteGroup>
    {
        /// <summary>
        /// Constructor for SiteGroupCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public SiteGroupCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
