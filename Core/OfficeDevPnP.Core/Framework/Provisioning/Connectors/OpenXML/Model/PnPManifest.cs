using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model
{
    /// <summary>
    /// Manifest of a PnP OpenXML file
    /// </summary>
    [Serializable]
    public class PnPManifest
    {
        /// <summary>
        /// The Type of the package file defined by the current manifest
        /// </summary>
        public PackageType Type { get; set; } = PackageType.Full;

        public String Version
        {
            get { return ("1.0"); }
        }
    }

    public enum PackageType
    {
        Full,
        Delta,
    }
}
