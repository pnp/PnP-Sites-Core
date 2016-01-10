using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Localization objects
    /// </summary>
    public partial class LocalizationCollection: ProvisioningTemplateCollection<Localization>
    {
        public LocalizationCollection(ProvisioningTemplate parentTemplate): 
            base(parentTemplate)
        {
        }
    }
}
