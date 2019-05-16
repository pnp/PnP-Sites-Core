using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Extensibility
{
    /// <summary>
    /// Defines the interface for a type that can pre-process a Provisioning Template
    /// </summary>
    /// <remarks>
    /// This interface is not yet implemented, and it is defined for future usage
    /// </remarks>
    public interface IProvisioningTemplateProcessor
    {
        /// <summary>
        /// Method to pre-process a Provisioning Template
        /// </summary>
        /// <param name="template">The source template</param>
        /// <returns>The resulting template</returns>
        ProvisioningTemplate PreProcessTemplate(ProvisioningTemplate template);

        /// <summary>
        /// Method to pre-process a Provisioning Template Hierarchy
        /// </summary>
        /// <param name="hierarchy">The source hierarchy</param>
        /// <returns>The resulting hierarchy</returns>
        ProvisioningHierarchy PreProcessHierarchy(ProvisioningHierarchy hierarchy);
    }
}
