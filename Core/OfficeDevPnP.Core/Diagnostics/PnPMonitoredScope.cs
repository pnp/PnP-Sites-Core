using OfficeDevPnP.Core.Diagnostics.Tree;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeDevPnP.Core.Diagnostics
{
    /// <summary>
    /// Class holds the methods for logging on PnP Monitoring
    /// </summary>
    public sealed class PnPMonitoredScope : TreeNode<PnPMonitoredScope>, IDisposable
    {
        [ThreadStatic]
        internal static PnPMonitoredScope TopScope;

        private Stopwatch _stopWatch;
        private string _name;
        private Guid _correlationId;
        private int _threadId;

        /// <summary>
        /// Constructor for PnPMonitoredScope class
        /// </summary>
        public PnPMonitoredScope()
        {
            Guid g = Guid.NewGuid();
            StartScope($"Unnamed Scope {g}");
        }

        internal int ThreadId
        {
            get
            {
                return this._threadId;
            }
        }
        /// <summary>
        /// Gets or sets the source name
        /// </summary>
        public string Name
        {
            get
            {
                return this._name;
            }
            set
            {
                this._name = string.IsNullOrEmpty(value) ? string.Empty : value;
            }
        }

        /// <summary>
        /// Constructor for PnPMonitoredScope class
        /// </summary>
        /// <param name="name">Source name</param>
        public PnPMonitoredScope(string name)
        {
            StartScope(name);
        }


        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.LogEntry.set_Message(System.String)")]
        private void StartScope(string name)
        {


            _threadId = Thread.CurrentThread.ManagedThreadId;
            _stopWatch = new Stopwatch();
            _name = name;
            _stopWatch.Start();
            _correlationId = Guid.NewGuid();

            if (PnPMonitoredScope.TopScope == null)
            {
                PnPMonitoredScope.TopScope = this;
            }
            if (TopScope != this)
            {
                var lastnode = TopScope.Descendants.Any() ? TopScope.Descendants.LastOrDefault() : TopScope;
                ((PnPMonitoredScope)lastnode).Children.Add(this);
            }
            LogDebug(CoreResources.PnPMonitoredScope_Code_execution_started);
        }

        private void EndScope()
        {
            _stopWatch.Stop();
            LogDebug(CoreResources.PnPMonitoredScope_Code_execution_ended, _stopWatch.ElapsedMilliseconds);

            Trace.Flush();
            if (TopScope == this)
            {
                TopScope = null;
            }
            Parent = null;
        }

        /// <summary>
        /// Gets Correlation Guid
        /// </summary>
        public Guid CorrelationId
        {
            get { return _correlationId; }
        }

        /// <summary>
        /// Logs Error
        /// </summary>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogError(string message, params object[] args)
        {
            Log.Error(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                ThreadId = _threadId
            });
        }

        /// <summary>
        /// Logs Error
        /// </summary>
        /// <param name="ex">Exception object</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogError(Exception ex, string message, params object[] args)
        {
            Log.Error(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                Exception = ex,
                ThreadId = _threadId
            });
        }

        /// <summary>
        /// Logs Information
        /// </summary>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogInfo(string message, params object[] args)
        {
            Log.Info(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                ThreadId = _threadId
            });
        }
        /// <summary>
        /// Logs Information 
        /// </summary>
        /// <param name="ex">Exception object</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogInfo(Exception ex, string message, params object[] args)
        {
            Log.Info(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                Exception = ex,
                ThreadId = _threadId
            });
        }


        /// <summary>
        /// Logs Warning
        /// </summary>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogWarning(string message, params object[] args)
        {
            Log.Warning(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                ThreadId = _threadId
            });
        }


        /// <summary>
        /// Logs Warning
        /// </summary>
        /// <param name="ex">Exception object</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogWarning(Exception ex, string message, params object[] args)
        {
            Log.Warning(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                Exception = ex,
                ThreadId = _threadId
            });

        }

        /// <summary>
        /// Debug Log
        /// </summary>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogDebug(string message, params object[] args)
        {
            Log.Debug(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                ThreadId = _threadId
            });
        }

        /// <summary>
        /// Debug Log
        /// </summary>
        /// <param name="ex">Exception object</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public void LogDebug(Exception ex, string message, params object[] args)
        {
            Log.Debug(new LogEntry()
            {
                CorrelationId = TopScope.CorrelationId,
                EllapsedMilliseconds = _stopWatch.ElapsedMilliseconds,
                Message = string.Format(message, args),
                Source = Name,
                Exception = ex,
                ThreadId = _threadId
            });
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        /// <summary>
        /// Implements Disposable pattern
        /// </summary>
        /// <param name="disposing">Boolean flag for disposing</param>
        void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    EndScope();

                }
                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~PnPMonitoredScope() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            Dispose(true);
        }
        #endregion

    }
}
