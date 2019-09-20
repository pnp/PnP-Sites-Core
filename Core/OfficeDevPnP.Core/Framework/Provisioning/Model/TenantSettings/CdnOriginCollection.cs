using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of CdnOrigin objects
    /// </summary>
    public partial class CdnOriginCollection : BaseProvisioningTemplateObjectCollection<CdnOrigin>
    {
        /// <summary>
        /// Constructor for CdnOriginCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public CdnOriginCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
