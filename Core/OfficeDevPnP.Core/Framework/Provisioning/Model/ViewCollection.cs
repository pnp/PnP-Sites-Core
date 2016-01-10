using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of View objects
    /// </summary>
    public partial class ViewCollection : ProvisioningTemplateCollection<View>
    {
        public ViewCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
