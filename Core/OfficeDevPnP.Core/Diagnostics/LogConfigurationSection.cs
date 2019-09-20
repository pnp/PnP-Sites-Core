using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    /// <summary>
    /// Class dealing with LogConfigurationTracing section
    /// </summary>
    public class LogConfigurationTracingSection : ConfigurationSection
    {
        /// <summary>
        /// Gets or sets "loglevel" config property
        /// </summary>
        [ConfigurationProperty("logLevel", DefaultValue = LogLevel.Information, IsRequired = false)]
        public LogLevel LogLevel
        {
            get
            {
                return (LogLevel)this["logLevel"];
            }
            set
            { this["logLevel"] = value; }
        }

        /// <summary>
        /// Gets or sets "logger" config property
        /// </summary>
        [ConfigurationProperty("logger")]
        public LogConfigurationTracingLoggerElement Logger
        {
            get
            {
                return (LogConfigurationTracingLoggerElement)this["logger"];
            }
            set
            {
                this["logger"] = value;
            }
        }
    }
    /// <summary>
    /// Class dealing with LogConfigurationTracing element
    /// </summary>
    public class LogConfigurationTracingLoggerElement : ConfigurationElement
    {
        /// <summary>
        /// Gets or sets "type" config property
        /// </summary>
        [ConfigurationProperty("type", DefaultValue = "false", IsRequired = true)]
        public string Type
        {
            get
            {
                return (string)this["type"];
            }
            set
            { this["type"] = value; }
        }
    }
}