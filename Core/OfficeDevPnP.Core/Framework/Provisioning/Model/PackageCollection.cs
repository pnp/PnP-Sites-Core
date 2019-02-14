using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of Package objects
    /// </summary>
    public partial class PackageCollection : BaseProvisioningTemplateObjectCollection<Package>
    {
        /// <summary>
        /// Constructor for PackagesCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public PackageCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
