using OfficeDevPnP.Core.Diagnostics.Tree;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeDevPnP.Core.Diagnostics
{

    public sealed class PnPMonitoredScope : TreeNode<PnPMonitoredScope>, IDisposable
    {
        [ThreadStatic]
        internal static PnPMonitoredScope TopScope;

        private Stopwatch _stopWatch;
        private string _name;
        private Guid _correlationId;
        private int _threadId;

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

        public Guid CorrelationId
        {
            get { return _correlationId; }
        }

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
