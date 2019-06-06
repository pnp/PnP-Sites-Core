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

        /// <summary>
        /// The version of the package file
        /// </summary>
        public String Version
        {
            get { return ("1.0"); }
        }
    }

    /// <summary>
    /// Types of package
    /// </summary>
    public enum PackageType
    {
        /// <summary>
        /// Full Package
        /// </summary>
        Full,
        /// <summary>
        /// Delta Package
        /// </summary>
        Delta,
    }
}
