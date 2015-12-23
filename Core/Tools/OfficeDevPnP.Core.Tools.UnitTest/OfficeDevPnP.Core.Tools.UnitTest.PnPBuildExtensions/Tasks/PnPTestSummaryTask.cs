using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.MD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Tasks
{
    public class PnPTestSummary: Task
    {
        [Required]
        public String TestResultsPath
        {
            get;
            set;
        }

        public override bool Execute()
        {
            try
            {
                Log.LogMessageFromText(String.Format("PnPTestSummaryTask: processing information from folder {0}", TestResultsPath), MessageImportance.Normal);

                Dictionary<string, string> parameters = new Dictionary<string, string>();
                parameters.Add("MDPath", TestResultsPath);

                TestManager testManager = new TestManager(parameters);
                testManager.GenerateMDSummaryReport();

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
