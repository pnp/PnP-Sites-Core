using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{

    public sealed class PnPMonitoredScope : IDisposable
    {
        private Stopwatch _stopWatch;
        private string _name;
        internal string LocalMachine = Environment.MachineName;
        private Guid _correlationId;

        public PnPMonitoredScope(string name)
        {
            _stopWatch = new Stopwatch();
            _name = name;
            _stopWatch.Start();
            _correlationId = Guid.NewGuid();

            LogInfo("Code execution started");
            Indent();
        }

        public void Indent()
        {
            Trace.Indent();
        }

        /// <summary>
        /// Decreases the current IndentLevel by one.
        /// </summary>
        public void Unindent()
        {
            Trace.Unindent();
        }

        public Guid CorrelationId
        {
            get { return _correlationId; }
        }

        public void LogError(string message, params object[] args)
        {
            var log = GetLogEntry(_name, message, args);
            Trace.TraceError(log);
            WriteLogToConsole(log);
        }

        public void LogInfo(string message, params object[] args)
        {
            var log = GetLogEntry(_name, message, args);
            Trace.TraceInformation(log);
            WriteLogToConsole(log);
        }

        public void LogWarning(string message, params object[] args)
        {
            var log = GetLogEntry(_name, message, args);
            Trace.TraceWarning(log);
            WriteLogToConsole(log);
        }
        private string GetLogEntry(string source, string message, params object[] args)
        {
            try
            {
                string msg = string.Empty;

                if (args == null || args.Length == 0)
                {
                    msg = message.Replace("{", "{{").Replace("}", "}}");
                }
                else
                {
                    msg = String.Format(CultureInfo.CurrentCulture, message, args);
                }

                string log = string.Format(CultureInfo.CurrentCulture, "{0} [[{1}]] {2} {3}ms {4}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), source, msg, _stopWatch.ElapsedMilliseconds, this.CorrelationId);
                return log;
            }
            catch (Exception e)
            {
                return string.Format("Error while generating log information, {0}", e);
            }
        }


        public void Dispose()
        {
            _stopWatch.Stop();
            Unindent();
            LogInfo("Code execution ended", _stopWatch.ElapsedMilliseconds);
            Trace.Flush();
        }

        [Conditional("DEBUG")]
        private void WriteLogToConsole(string value)
        {
            var part1 = value.Substring(0, value.IndexOf("]]") + 2);
            var part2 = value.Substring(value.IndexOf("]]") + 2);
            Console.WriteLine("{0} {1}{2}", part1, new string(' ', Trace.IndentLevel * 2), part2);

        }
    }
}
