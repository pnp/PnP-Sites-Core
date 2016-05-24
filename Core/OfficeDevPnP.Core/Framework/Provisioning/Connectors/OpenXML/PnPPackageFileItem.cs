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
        public String Name { get; set; }

        public String Folder { get; set; }

        public Byte[] Content { get; set; }
    }
}
