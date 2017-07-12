using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    /// <summary>
    /// Interface for Logging
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// Log Information
        /// </summary>
        /// <param name="entry">LogEntry object</param>
        void Info(LogEntry entry);
        /// <summary>
        /// Warning Log
        /// </summary>
        /// <param name="entry">LogEntry object</param>
        void Warning(LogEntry entry);
        /// <summary>
        /// Error Log
        /// </summary>
        /// <param name="entry">LogEntry object</param>
        void Error(LogEntry entry);
        /// <summary>
        /// Debug Log
        /// </summary>
        /// <param name="entry">LogEntry object</param>
        void Debug(LogEntry entry);
    }
}
