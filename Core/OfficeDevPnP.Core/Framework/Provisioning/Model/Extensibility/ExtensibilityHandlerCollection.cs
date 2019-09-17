using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of ExtensibilityHandler objects
    /// </summary>
    public partial class ExtensibilityHandlerCollection : BaseProvisioningTemplateObjectCollection<ExtensibilityHandler>
    {
        /// <summary>
        /// Constructor for ExtensibilityHandlerCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ExtensibilityHandlerCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
