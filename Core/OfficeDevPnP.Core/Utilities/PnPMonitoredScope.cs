using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
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

        public PnPMonitoredScope()
        {
            StackFrame frame = new StackFrame(1);
            var method = frame.GetMethod();
            var name = method.Name;
            StartScope(name);
        }

        public PnPMonitoredScope(string name)
        {
            StartScope(name);
        }

        private void StartScope(string name)
        {
            _stopWatch = new Stopwatch();
            _name = name;
            _stopWatch.Start();
            _correlationId = Guid.NewGuid();

            LogInfo(CoreResources.PnPMonitoredScope_Code_execution_started);
            Indent();
        }

        private void EndScope()
        {
            _stopWatch.Stop();
            Unindent();
            LogInfo(CoreResources.PnPMonitoredScope_Code_execution_ended, _stopWatch.ElapsedMilliseconds);
            Trace.Flush();
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

                string log = string.Format(CultureInfo.CurrentCulture, "{0} [{1}] {2} {3}ms {4}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), source, msg, _stopWatch.ElapsedMilliseconds, this.CorrelationId);
                return log;
            }
            catch (Exception e)
            {
                return string.Format("Error while generating log information, {0}", e);
            }
        }


        public void Dispose()
        {
            EndScope();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Console.WriteLine(System.String,System.Object,System.Object)")]
        [Conditional("DEBUG")]
        private void WriteLogToConsole(string value)
        {
            Console.WriteLine("{0}{1}", new string(' ', Trace.IndentLevel * 2), value);
        }
    }
}
