using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of ProvisioningSequence items
    /// </summary>
    public partial class ProvisioningSequenceCollection : BaseProvisioningHierarchyObjectCollection<ProvisioningSequence>
    {
        /// <summary>
        /// Constructor for ProvisioningSequenceCollection class
        /// </summary>
        /// <param name="parentProvisioning">Parent Provisioning object</param>
        public ProvisioningSequenceCollection(ProvisioningHierarchy parentProvisioning) :
            base(parentProvisioning)
        {
        }
    }
}
