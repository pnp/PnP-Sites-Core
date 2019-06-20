using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of AddIn objects
    /// </summary>
    public partial class AppCollection : BaseProvisioningTemplateObjectCollection<App>
    {
        /// <summary>
        /// Constructor for AppCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public AppCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
