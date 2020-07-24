using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Interface implemented by any descendant of a ProvisioningTemplate
    /// </summary>
    public interface IProvisioningTemplateDescendant
    {
        /// <summary>
        /// References the parent ProvisioningTemplate for the current provisioning artifact
        /// </summary>
        ProvisioningTemplate ParentTemplate { get; }
    }
}
