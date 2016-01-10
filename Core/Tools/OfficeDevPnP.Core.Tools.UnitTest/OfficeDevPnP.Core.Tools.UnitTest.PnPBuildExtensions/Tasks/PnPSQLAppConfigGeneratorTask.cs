using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Tasks
{
    public class PnPSQLAppConfigGeneratorTask: Task
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

        [Required]
        public String AppConfigFolder
        {
            get;
            set;
        }


        public override bool Execute()
        {
            try
            {
                Log.LogMessageFromText(String.Format("PnPSqlAppConfigGeneratorTask: Reading information for configuration {0} to generate app.config in {1}", Configuration, AppConfigFolder), MessageImportance.Normal);
                PnPAppConfigManager appConfigManager = new PnPAppConfigManager(SQLConnectionString.Replace("&quot;", "\""), Configuration);
                appConfigManager.GenerateAppConfig(AppConfigFolder);
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

