using System;
using System.Threading;

namespace OfficeDevPnP.Core.Diagnostics
{
    /// <summary>
    /// Logging class
    /// </summary>
    public static class Log
    {
        [ThreadStatic]
        private static ILogger _logger;

        [ThreadStatic]
        private static LogLevel? _logLevel;
        /// <summary>
        /// Gets or sets Log Level
        /// </summary>
        public static LogLevel LogLevel
        {
            get { return _logLevel.Value; }
            set { _logLevel = value; }
        }

        /// <summary>
        /// Gets or sets ILogger object
        /// </summary>
        public static ILogger Logger
        {
            get { return _logger; }
            set { _logger = value; }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.LogEntry.set_Message(System.String)")]
        private static void InitializeLogger()
        {
            if (_logger == null)
            {
                var config = (OfficeDevPnP.Core.Diagnostics.LogConfigurationTracingSection)System.Configuration.ConfigurationManager.GetSection("pnp/tracing");

                if (config != null)
                {
                    _logLevel = config.LogLevel;

                    try
                    {
                        if (config.Logger.ElementInformation.IsPresent)
                        {
                            var loggerType = Type.GetType(config.Logger.Type, false);
#if !NETSTANDARD2_0
                            _logger = (ILogger)Activator.CreateInstance(loggerType.Assembly.FullName, loggerType.FullName).Unwrap();
#else
                            _logger = (ILogger)Activator.CreateInstance(loggerType);
#endif
                        }
                        else
                        {
                            _logger = new TraceLogger();
                        }
                    }
                    catch (Exception ex)
                    {
                        // Something went wrong, fall back to the built-in PnPTraceLogger
                        _logger = new TraceLogger();
                        _logger.Error(
                            new LogEntry()
                            {
                                Exception = ex,
                                Message = "Logger registration failed. Falling back to TraceLogger.",
                                EllapsedMilliseconds = 0,
                                CorrelationId = Guid.Empty,
                                ThreadId = Thread.CurrentThread.ManagedThreadId,
                                Source = "PnP"
                            });
                    }
                }
                else
                {
                    // Defaulting to built in logger
                    if (!_logLevel.HasValue)
                    {
                        _logLevel = LogLevel.Debug;
                    }
                    _logger = new TraceLogger();
                }
            }
        }

#region Public Members

#region Error
        /// <summary>
        /// Logs error message and source
        /// </summary>
        /// <param name="source">Error source</param>
        /// <param name="message">Error message</param>
        /// <param name="args">Arguments object</param>
        public static void Error(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Error || _logLevel == LogLevel.Debug)
            {
                _logger.Error(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source
                });
            }
        }
        /// <summary>
        /// Logs error message, source and exception
        /// </summary>
        /// <param name="ex">Exception object</param>
        /// <param name="source">Error source</param>
        /// <param name="message">Error message</param>
        /// <param name="args">Arguments object</param>
        public static void Error(Exception ex, string source, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Error || _logLevel == LogLevel.Debug)
            {
                _logger.Error(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }
        /// <summary>
        /// Error LogEntry
        /// </summary>
        /// <param name="logEntry">LogEntry object</param>
        public static void Error(LogEntry logEntry)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Error || _logLevel == LogLevel.Debug)
            {
                _logger.Error(logEntry);
            }
        }
#endregion

#region Info
        /// <summary>
        /// Log Information
        /// </summary>
        /// <param name="source">Source string</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public static void Info(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Information || _logLevel == LogLevel.Debug || _logLevel == LogLevel.Error || _logLevel == LogLevel.Warning)
            {
                _logger.Info(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source
                });
            }
        }
        /// <summary>
        /// Log Information
        /// </summary>
        /// <param name="ex">Exception object</param>
        /// <param name="source">Source string</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments option</param>
        public static void Info(Exception ex, string source, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Information || _logLevel == LogLevel.Debug || _logLevel == LogLevel.Error || _logLevel == LogLevel.Warning)
            {
                _logger.Info(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }
        /// <summary>
        /// Log Information
        /// </summary>
        /// <param name="logEntry">LogEntry object</param>
        public static void Info(LogEntry logEntry)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Information || _logLevel == LogLevel.Debug || _logLevel == LogLevel.Error || _logLevel == LogLevel.Warning)
            {
                _logger.Info(logEntry);
            }
        }
#endregion

#region Warning
        /// <summary>
        /// Warning Log
        /// </summary>
        /// <param name="source">Source string</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public static void Warning(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Warning || _logLevel == LogLevel.Information || _logLevel == LogLevel.Debug)
            {
                _logger.Warning(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                });
            }
        }
        /// <summary>
        /// Warning Log
        /// </summary>
        /// <param name="source">Source string</param>
        /// <param name="ex">Exception object</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public static void Warning(string source, Exception ex, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Warning || _logLevel == LogLevel.Information || _logLevel == LogLevel.Debug)
            {
                _logger.Warning(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }

        /// <summary>
        /// Warning Log
        /// </summary>
        /// <param name="logEntry">LogEntry object</param>
        public static void Warning(LogEntry logEntry)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Warning || _logLevel == LogLevel.Information || _logLevel == LogLevel.Debug)
            {
                _logger.Warning(logEntry);
            }
        }
#endregion

#region Debug
        /// <summary>
        /// Debug Log
        /// </summary>
        /// <param name="source">Source stirng</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public static void Debug(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Debug)
            {
                _logger.Debug(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                });
            }
        }


        /// <summary>
        /// Debug Log
        /// </summary>
        /// <param name="source">Source string</param>
        /// <param name="ex">Exception object</param>
        /// <param name="message">Message string</param>
        /// <param name="args">Arguments object</param>
        public static void Debug(string source, Exception ex, string message, params object[] args)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Debug)
            {
                _logger.Debug(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }

        /// <summary>
        /// Debug Log
        /// </summary>
        /// <param name="logEntry">LogEntry object</param>
        public static void Debug(LogEntry logEntry)
        {
            InitializeLogger();
            if (_logLevel == LogLevel.Debug)
            {
                _logger.Debug(logEntry);
            }
        }
#endregion

#endregion
    }
}