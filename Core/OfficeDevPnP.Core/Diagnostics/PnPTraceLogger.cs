using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    class PnPTraceLogger : ILogger
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.LogEntry.set_Message(System.String)")]
        public void Debug(LogEntry entry)
        {
            entry.Message = "[DEBUG] " + entry.Message;
            Trace.TraceInformation(GetLogEntry(entry));
        }

        public void Error(LogEntry entry)
        {
            Trace.TraceError(GetLogEntry(entry));
        }

        public void Info(LogEntry entry)
        {
            Trace.TraceInformation(GetLogEntry(entry));
        }

        public void TraceApi(LogEntry entry)
        {
            // not implemented
        }

        public void Warning(LogEntry entry)
        {
            Trace.TraceWarning(GetLogEntry(entry));
        }

        private string GetLogEntry(LogEntry entry)
        {

            try
            {
                string log = string.Format("{0}\t[{1}]:[{2}]\t{3}\t{4}ms\t{5}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), entry.Source, entry.ThreadId, entry.Message, entry.EllapsedMilliseconds, entry.CorrelationId);

                return log;
            }
            catch (Exception e)
            {
                return string.Format("Error while generating log information, {0}", e);
            }
        }
    }
}
