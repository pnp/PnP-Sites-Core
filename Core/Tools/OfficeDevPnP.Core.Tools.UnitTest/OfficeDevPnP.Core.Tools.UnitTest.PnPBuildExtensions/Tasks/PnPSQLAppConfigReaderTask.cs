using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Tasks
{
    public class PnPSQLAppConfigReaderTask : Task
    {
        [Required]
        public string SQLConnectionString
        {
            get;
            set;
        }

        [Required]
        public String Configuration
        {
            get;
            set;
        }

        [Output]
        public string PnPBuildConfiguration
        {
            get;
            set;
        }

        [Output]
        public string PnPBranch
        {
            get;
            set;
        }

        public override bool Execute()
        {
            try
            {
                //Log.LogMessageFromText(String.Format("PnPAppConfigReaderTask: Reading information for configuration {0}", Configuration), MessageImportance.Normal);
                PnPAppConfigManager appConfigManager = new PnPAppConfigManager(SQLConnectionString.Replace("&quot;", "\""), Configuration);
                PnPBuildConfiguration = appConfigManager.GetConfigurationElement("PnPBuild");
                PnPBranch = appConfigManager.GetConfigurationElement("PnPBranch");
                return true;
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
                return false;
            }
        }

    }
}
