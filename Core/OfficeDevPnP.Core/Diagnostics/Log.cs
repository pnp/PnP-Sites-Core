using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    /// <summary>
    /// Logging class
    /// </summary>
    public static class Log
    {
        [ThreadStatic]
        private static ILogger logger;

        [ThreadStatic]
        private static LogLevel logLevel;

        public static LogLevel LogLevel
        {
            get { return logLevel; }
            set { logLevel = value; }
        }


        private static void InitializeLogger()
        {
            if (logger == null)
            {
                var config = (OfficeDevPnP.Core.Diagnostics.LogConfigurationTracingSection)System.Configuration.ConfigurationManager.GetSection("pnp/tracing");

                if (config != null)
                {
                    logLevel = config.LogLevel;

                    try
                    {
                        if (config.Logger.ElementInformation.IsPresent)
                        {
                            logger = (ILogger)Activator.CreateInstance(config.Logger.Assembly, config.Logger.Type).Unwrap();
                        }
                        else
                        {
                            logger = new PnPTraceLogger();
                        }
                    }
                    catch (Exception ex)
                    {
                        // Something went wrong, fall back to the built-in PnPTraceLogger
                        logger = new PnPTraceLogger();
                        logger.Error(
                            new LogEntry()
                            {
                                Exception = ex,
                                Message = "Logger registration failed. Falling back to PnPTraceLogger.",
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
                    logLevel = LogLevel.Information;
                    logger = new PnPTraceLogger();
                }
            }
        }

        #region Public Members

        #region Error
        public static void Error(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Error || logLevel == LogLevel.Debug)
            {
                logger.Error(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source
                });
            }
        }


        public static void Error(Exception ex, string source, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Error || logLevel == LogLevel.Debug)
            {
                logger.Info(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }

        public static void Error(LogEntry logEntry)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Error || logLevel == LogLevel.Debug)
            {
                logger.Error(logEntry);
            }
        }
        #endregion

        #region Info
        public static void Info(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Information || logLevel == LogLevel.Debug || logLevel == LogLevel.Error || logLevel == LogLevel.Warning)
            {
                logger.Info(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source
                });
            }
        }


        public static void Info(Exception ex, string source, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Information || logLevel == LogLevel.Debug || logLevel == LogLevel.Error || logLevel == LogLevel.Warning)
            {
                logger.Info(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }

        public static void Info(LogEntry logEntry)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Information || logLevel == LogLevel.Debug || logLevel == LogLevel.Error || logLevel == LogLevel.Warning)
            {
                logger.Info(logEntry);
            }
        }
        #endregion

        #region Warning

        public static void Warning(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Warning || logLevel == LogLevel.Information)
            {
                logger.Warning(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                });
            }
        }



        public static void Warning(string source, Exception ex, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Warning || logLevel == LogLevel.Information)
            {
                logger.Warning(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }

        public static void Warning(LogEntry logEntry)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Debug)
            {
                logger.Warning(logEntry);
            }
        }
        #endregion

        #region Debug
        public static void Debug(string source, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Debug)
            {
                logger.Debug(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                });
            }
        }



        public static void Debug(string source, Exception ex, string message, params object[] args)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Debug)
            {
                logger.Debug(new LogEntry()
                {
                    Message = string.Format(message, args),
                    Source = source,
                    Exception = ex,
                });
            }
        }

        public static void Debug(LogEntry logEntry)
        {
            InitializeLogger();
            if (logLevel == LogLevel.Debug)
            {
                logger.Debug(logEntry);
            }
        }
        #endregion

        #endregion
    }
}
