using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL
{
    [ExtensionUri("logger://SQLLogger/v1")] /// Uri used to uniquely identify the SQL logger. 
    [FriendlyName("SQLLogger")] /// Alternate user friendly string to uniquely identify the logger.
    public class VSTestLogger : ITestLoggerWithParameters
    {
        private TestManager testManager;

        public void Initialize(TestLoggerEvents events, string testRunDirectory)
        {
            //Console.WriteLine("In Initialize");
        }

        public void Initialize(TestLoggerEvents events, Dictionary<string, string> parameters)
        {
            //foreach (var param in parameters)
            //{
            //    Console.WriteLine("Property: {0}   Value:{1}", param.Key, param.Value);
            //}

            testManager = new TestManager(parameters);

            events.TestResult += Events_TestResult;
            events.TestRunMessage += Events_TestRunMessage;
            events.TestRunComplete += Events_TestRunComplete;
        }

        private void Events_TestResult(object sender, Microsoft.VisualStudio.TestPlatform.ObjectModel.Logging.TestResultEventArgs e)
        {
            testManager.AddTestResult(e.Result);
        }

        private void Events_TestRunComplete(object sender, TestRunCompleteEventArgs e)
        {
            try
            {
                testManager.TestAreDone(e.TestRunStatistics, e.IsCanceled, e.IsAborted, e.Error, e.AttachmentSets, e.ElapsedTimeInRunningTests);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: {0}", ex.ToString());
            }
        }

        private void Events_TestRunMessage(object sender, Microsoft.VisualStudio.TestPlatform.ObjectModel.Logging.TestRunMessageEventArgs e)
        {
            Console.WriteLine(e.Message);
        }

    }
}
