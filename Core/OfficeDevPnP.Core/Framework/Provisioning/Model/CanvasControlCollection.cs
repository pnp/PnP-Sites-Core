using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of CanvasControl objects
    /// </summary>
    public partial class CanvasControlCollection : BaseProvisioningTemplateObjectCollection<CanvasControl>
    {
        /// <summary>
        /// Constructor for CanvasControlCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public CanvasControlCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
