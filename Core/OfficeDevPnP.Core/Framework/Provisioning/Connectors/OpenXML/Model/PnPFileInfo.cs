using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model
{
    /// <summary>
    /// File descriptor for every single file in the PnP OpenXML file
    /// </summary>
    [Serializable]
    public class PnPFileInfo
    {
        /// <summary>
        /// The Name of the file in the PnP OpenXML file
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// The name of the folder within the PnP OpenXML file
        /// </summary>
        public String Folder { get; set; }

        /// <summary>
        /// The binary content of the file
        /// </summary>
        public Byte[] Content { get; set; }
    }
}
