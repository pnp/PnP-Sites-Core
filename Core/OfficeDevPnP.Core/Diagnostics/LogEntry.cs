using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    public class LogEntry
    {
        public string Message { get; set; }
        public Guid CorrelationId { get; set; }
        public string Source { get; set; }
        public Exception Exception { get; set; }
        public int ThreadId { get; set; }
        public long EllapsedMilliseconds { get; set; }
    }
}
