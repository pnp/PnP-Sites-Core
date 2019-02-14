using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Interface implemented by a descendant of a Provisioning object
    /// </summary>
    public interface IProvisioningHierarchyDescendant
    {
        /// <summary>
        /// References the parent Provisioning for the current artifact
        /// </summary>
        ProvisioningHierarchy ParentHierarchy { get; }
    }
}
