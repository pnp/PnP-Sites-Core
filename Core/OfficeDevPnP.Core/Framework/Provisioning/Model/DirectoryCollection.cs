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
    public partial class DirectoryCollection : BaseProvisioningTemplateObjectCollection<Directory>
    {
        /// <summary>
        /// Constructor for DirectoryCollection class.
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public DirectoryCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
