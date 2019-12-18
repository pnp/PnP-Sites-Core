using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Base type for any Domain Model object in the Provisioning Hierarchy (from the ProvisioningTemplate type and above)
    /// </summary>
    public abstract class BaseHierarchyModel : IProvisioningHierarchyDescendant
    {
        private ProvisioningHierarchy _parentHierarchy;

        /// <summary>
        /// Represents a reference to the parent Provisioning Hierarchy object, if any
        /// </summary>
        /// <remarks>
        /// Introduced to support schema v2018-07 and tenant level provisioning
        /// </remarks>
        [JsonIgnore]
        public ProvisioningHierarchy ParentHierarchy
        {
            get
            {
                return (this._parentHierarchy);
            }
            internal set
            {
                this._parentHierarchy = value;
            }
        }
    }
}
