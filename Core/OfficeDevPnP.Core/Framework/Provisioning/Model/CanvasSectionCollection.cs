using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of CanvasZone objects
    /// </summary>
    public partial class CanvasSectionCollection : BaseProvisioningTemplateObjectCollection<CanvasSection>
    {
        /// <summary>
        /// Constructor for CanvasZoneCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public CanvasSectionCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
