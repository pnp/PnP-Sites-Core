using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Drive
{
    /// <summary>
    /// Collection of DriveItem items
    /// </summary>
    public class DriveItemCollection : BaseProvisioningTemplateObjectCollection<DriveItem>
    {
        /// <summary>
        /// Constructor for DriveItemCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public DriveItemCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
