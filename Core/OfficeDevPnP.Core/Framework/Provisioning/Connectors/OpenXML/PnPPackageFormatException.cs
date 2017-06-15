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
        /// <summary>
        /// Constructor for PnPPackageFormatException class with the specified error message.
        /// </summary>
        /// <param name="message">A string that describes the exception</param>
        public PnPPackageFormatException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Constructor for PnPackageFormatException class with the specified error message and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">A string that describes the exception</param>
        /// <param name="innerException">The exception that is the cause of the current exception</param>
        public PnPPackageFormatException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
