using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of SiteCollection items
    /// </summary>
    public class SiteCollectionCollection : BaseProvisioningObjectCollection<SiteCollection>
    {
        /// <summary>
        /// Constructor for SiteCollectionCollection class
        /// </summary>
        /// <param name="parentProvisioning">Parent Provisioning object</param>
        public SiteCollectionCollection(Provisioning parentProvisioning) :
            base(parentProvisioning)
        {
        }
    }
}
