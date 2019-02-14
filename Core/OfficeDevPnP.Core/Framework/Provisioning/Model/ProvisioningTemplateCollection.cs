using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of ProvisioningTemplate items
    /// </summary>
    public class ProvisioningTemplateCollection : BaseProvisioningHierarchyObjectCollection<ProvisioningTemplate>
    {
        /// <summary>
        /// Constructor for ProvisioningTemplateCollection class
        /// </summary>
        /// <param name="parentProvisioning">Parent Provisioning object</param>
        public ProvisioningTemplateCollection(ProvisioningHierarchy parentProvisioning) :
            base(parentProvisioning)
        {
        }
    }
}
