using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Provides the complex output of the CanProvision method
    /// </summary>
    public class CanProvisionResult
    {
        /// <summary>
        /// Defines whether the Provisioning can proceed or not
        /// </summary>
        public Boolean CanProvision { get; set; }

        /// <summary>
        /// The list of detailed CanProvision Issues, if any
        /// </summary>
        public List<CanProvisionIssue> Issues { get; set; }
    }

    /// <summary>
    /// Defines a CanProvision Issue item
    /// </summary>
    public class CanProvisionIssue
    {
        /// <summary>
        /// The Source of the CanProvision Issue
        /// </summary>
        public String Source { get; set; }

        /// <summary>
        /// Provides a text-based description of the Issue
        /// </summary>
        public String Message { get; set; }

        /// <summary>
        /// Provides a unique Tag for the current issue
        /// </summary>
        public Int32 Tag { get; set; }

        /// <summary>
        /// Provides the Inner Exception for a blocking (Severity = Error) Issue
        /// </summary>
        public Exception InnerException { get; set; }
    }
}
