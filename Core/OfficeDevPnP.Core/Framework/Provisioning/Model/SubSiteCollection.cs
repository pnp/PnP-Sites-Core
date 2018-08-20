using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of SubSite items
    /// </summary>
    public class SubSiteCollection : BaseProvisioningObjectCollection<SubSite>
    {
        /// <summary>
        /// Constructor for SubSiteCollection class
        /// </summary>
        /// <param name="parentProvisioning">Parent Provisioning object</param>
        public SubSiteCollection(Provisioning parentProvisioning) :
            base(parentProvisioning)
        {
        }
    }
}
