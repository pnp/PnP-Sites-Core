using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    class TraceLogger : ILogger
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.LogEntry.set_Message(System.String)")]
        public void Debug(LogEntry entry)
        {
            entry.Message = entry.Message;
            Trace.TraceInformation(GetLogEntry(entry, LogLevel.Debug));
        }

        public void Error(LogEntry entry)
        {
            Trace.TraceError(GetLogEntry(entry, LogLevel.Error));
        }

        public void Info(LogEntry entry)
        {
            Trace.TraceInformation(GetLogEntry(entry, LogLevel.Information));
        }

        public void Warning(LogEntry entry)
        {
            Trace.TraceWarning(GetLogEntry(entry, LogLevel.Information));
        }

        private string GetLogEntry(LogEntry entry, LogLevel level)
        {
            try
            {
                string log = string.Format("{0}\t[{1}]\t[{2}]\t[{3}]\t{4}\t{5}ms\t{6}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff"), entry.Source, entry.ThreadId, level.ToString(), entry.Message, entry.EllapsedMilliseconds, entry.CorrelationId != Guid.Empty ? entry.CorrelationId.ToString() : "");

                return log;
            }
            catch (Exception e)
            {
                return $"Error while generating log information, {e}";
            }
        }
    }
}
