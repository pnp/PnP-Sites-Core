using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of AlternateUICulture objects
    /// </summary>
    public partial class AlternateUICultureCollection : BaseProvisioningTemplateObjectCollection<AlternateUICulture>
    {
        /// <summary>
        /// Constructor for AlternateUICultureCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public AlternateUICultureCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
