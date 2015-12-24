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
    public class PnPSQLStartTestTask: Task
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
        public int PnPTestRunId
        {
            get;
            set;
        }

        public override bool Execute()
        {
            try
            {
                //System.Diagnostics.Debugger.Launch();

                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters.Add("PnPSQLConnectionString", PnPBase64EncoderTask.Base64Encode(SQLConnectionString).Replace("=", "&equal"));
                parameters.Add("PnPConfigurationToTest", Configuration);

                TestManager testManager = new TestManager(parameters);
                PnPTestRunId = testManager.AddTestSetRecord();
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
