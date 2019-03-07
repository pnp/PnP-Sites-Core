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
        /// Declares the Severity of the Issue
        /// </summary>
        public CanProvisionIssueSeverity Severity { get; set; }

        /// <summary>
        /// Provides a text-based description of the Issue
        /// </summary>
        public String Message { get; set; }

        /// <summary>
        /// Provides the Inner Exception for a blocking (Severity = Error) Issue
        /// </summary>
        public Exception InnerException { get; set; }
    }

    /// <summary>
    /// Defines the Severity of a CanProvision Issue
    /// </summary>
    public enum CanProvisionIssueSeverity
    {
        /// <summary>
        /// The CanProvision Issue is not blocking, it is just an informative issue, the provisioning can proceed
        /// </summary>
        Information,
        /// <summary>
        /// The CanProvision Issue is not blocking, but the provisioning can proceed upon user's confirmation
        /// </summary>
        Warning,
        /// <summary>
        /// The CanProvision Issue is blocking, the provisioning cannot proceed
        /// </summary>
        Error,
    }
}
