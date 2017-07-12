using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Folder objects
    /// </summary>
    public partial class FolderCollection : ProvisioningTemplateCollection<Folder>
    {
        /// <summary>
        /// Constructor for Folder class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public FolderCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
