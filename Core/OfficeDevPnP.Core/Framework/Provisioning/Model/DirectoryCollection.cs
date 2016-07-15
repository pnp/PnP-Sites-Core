using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Directory objects
    /// </summary>
    public partial class DirectoryCollection : ProvisioningTemplateCollection<Directory>
    {
        public DirectoryCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
