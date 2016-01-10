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
    public abstract class BaseModel : IProvisioningTemplateDescendant
    {
        private ProvisioningTemplate _parentTemplate;

        /// <summary>
        /// References the parent ProvisioningTemplate for the current provisioning artifact
        /// </summary>
        [JsonIgnore]
        public ProvisioningTemplate ParentTemplate
        {
            get
            {
                return (this._parentTemplate);
            }
            internal set
            {
                this._parentTemplate = value;
            }
        }
    }
}
