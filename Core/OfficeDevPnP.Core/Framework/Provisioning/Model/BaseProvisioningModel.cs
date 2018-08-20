using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Base type for any Domain Model object (excluded the ProvisioningTemplate type)
    /// </summary>
    public abstract class BaseProvisioningModel : IProvisioningDescendant
    {
        private Provisioning _parentProvisioning;

        /// <summary>
        /// Represents a reference to the parent Provisioning object, if any
        /// </summary>
        /// <remarks>
        /// Introduced to support schema v2018-07 and tenant level provisioning
        /// </remarks>
        [JsonIgnore]
        public Provisioning ParentProvisioning
        {
            get
            {
                return (this._parentProvisioning);
            }
            internal set
            {
                this._parentProvisioning = value;
            }
        }
    }
}
