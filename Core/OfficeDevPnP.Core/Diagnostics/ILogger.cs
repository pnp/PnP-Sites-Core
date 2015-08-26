using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    public interface ILogger
    {
        void Info(LogEntry entry);
        void Warning(LogEntry entry);
        void Error(LogEntry entry);
        void Debug(LogEntry entry);
    }
}
