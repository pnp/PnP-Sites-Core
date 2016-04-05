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
    public partial class ExtensibilityHandlerCollection : ProvisioningTemplateCollection<ExtensibilityHandler>
    {
        public ExtensibilityHandlerCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
