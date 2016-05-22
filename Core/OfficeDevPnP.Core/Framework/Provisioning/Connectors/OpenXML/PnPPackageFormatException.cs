using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML
{
    /// <summary>
    /// Custom Exception type for PnP Packaging handling
    /// </summary>
    public class PnPPackageFormatException : ApplicationException
    {
        public PnPPackageFormatException(string message)
            : base(message)
        {
        }

        public PnPPackageFormatException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
