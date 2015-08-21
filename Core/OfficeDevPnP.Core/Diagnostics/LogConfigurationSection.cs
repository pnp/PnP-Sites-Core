using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics
{
    public class LogConfigurationTracingSection : ConfigurationSection
    {
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

    public class LogConfigurationTracingLoggerElement : ConfigurationElement
    {
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
