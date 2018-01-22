using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML
{
    /// <summary>
    /// Defines a single file in the PnP Open XML file package
    /// </summary>
    public class PnPPackageFileItem
    {
        /// <summary>
        /// Name of the package file item
        /// </summary>
        public String Name { get; set; }
        /// <summary>
        /// Folder containing the package file item
        /// </summary>
        public String Folder { get; set; }
        /// <summary>
        /// Content of the package file item
        /// </summary>
        public Byte[] Content { get; set; }
    }
}
