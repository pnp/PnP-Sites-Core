using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    /// <summary>
    /// Class holds LogEntry properties
    /// </summary>
    public class LogEntry
    {
        /// <summary>
        /// Gets or sets Log message
        /// </summary>
        public string Message { get; set; }
        /// <summary>
        /// Gets or sets CorrelationId of type Guid
        /// </summary>
        public Guid CorrelationId { get; set; }
        /// <summary>
        /// Gets or sets Log source
        /// </summary>
        public string Source { get; set; }
        /// <summary>
        /// Gets or sets Log Exception
        /// </summary>
        public Exception Exception { get; set; }
        /// <summary>
        /// Gets or sets Log ThreadId
        /// </summary>
        public int ThreadId { get; set; }
        /// <summary>
        /// Gets or sets elapsed Log time in MilliSeconds
        /// </summary>
        public long EllapsedMilliseconds { get; set; }
    }
}
