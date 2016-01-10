using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Entity.Validation;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL
{
    public class TestManager
    {
        #region Private variables
        private Dictionary<string, string> loggerParameters;
        private TestModelContainer context;
        private TestConfiguration testConfiguration;
        private TestRun testRun;
        private bool firstTest = true;
        #endregion

        #region Constructor
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="parameters">Dictionary with parameter values</param>
        public TestManager(Dictionary<string, string> parameters)
        {
            //System.Diagnostics.Debugger.Launch();
            loggerParameters = parameters;

            // validate is we've the needed params
            if (String.IsNullOrEmpty(GetParameter("PnPSQLConnectionString")) || String.IsNullOrEmpty(GetParameter("PnPConfigurationToTest")))
            {
                throw new ArgumentException("Requested parameters (PnPSQLConnectionString and PnPConfigurationToTest) are not defined");
            }

            // we pass the connection string as base64 encoded + replaced "=" with &quot; to avoid problems with the default implementation of the VSTestLogger interface
            context = new TestModelContainer(Base64Decode(GetParameter("PnPSQLConnectionString")).Replace("&quot;", "\""));

            // find the used configuration
            string configurationToTest = GetParameter("PnPConfigurationToTest");
            testConfiguration = context.TestConfigurationSet.Where(s => s.Name.Equals(configurationToTest, StringComparison.InvariantCultureIgnoreCase)).First();
            if (testConfiguration == null)
            {
                throw new Exception(String.Format("Test configuration with name {0} was not found", configurationToTest));
            }
        }
        #endregion

        #region public methods
        public int AddTestSetRecord()
        {
            // Grab the build of the environment we're testing
            string build = GetBuildNumber();

            // Log a record to indicate we're starting up the testing
            testRun = new TestRun()
            {
                TestDate = DateTime.Now,
                Build = build,
                Status = RunStatus.Initializing,
                TestWasAborted = false,
                TestWasCancelled = false,
                TestConfiguration = testConfiguration,
            };
            context.TestRunSet.Add(testRun);

            // persist to the database
            SaveChanges();

            return testRun.Id;
        }

        /// <summary>
        /// Adds a test result to the database
        /// </summary>
        /// <param name="test">Test result</param>
        public void AddTestResult(Microsoft.VisualStudio.TestPlatform.ObjectModel.TestResult test)
        {
            if (firstTest)
            {
                firstTest = false;
                // If the testRun record was already created then grab it else create a new one
                int testRunId;
                if (!String.IsNullOrEmpty(GetParameter("PnPTestRunId")) && Int32.TryParse(GetParameter("PnPTestRunId"), out testRunId))
                {
                    testRun = context.TestRunSet.Find(testRunId);
                }
                else
                {
                    AddTestSetRecord();
                }

                // Bring status to "running"
                testRun.Status = RunStatus.Running;
            }

            // Store the test result
            TestResult tr = new TestResult()
            {
                ComputerName = test.ComputerName,
                TestCaseName = !string.IsNullOrEmpty(test.DisplayName) ? test.DisplayName : test.TestCase.FullyQualifiedName,
                Duration = test.Duration,
                ErrorMessage = test.ErrorMessage,
                ErrorStackTrace = test.ErrorStackTrace,
                StartTime = test.StartTime,
                EndTime = test.EndTime,
            };

            switch (test.Outcome)
            {
                case Microsoft.VisualStudio.TestPlatform.ObjectModel.TestOutcome.None:
                    tr.Outcome = Outcome.None;
                    break;
                case Microsoft.VisualStudio.TestPlatform.ObjectModel.TestOutcome.Passed:
                    tr.Outcome = Outcome.Passed;
                    break;
                case Microsoft.VisualStudio.TestPlatform.ObjectModel.TestOutcome.Failed:
                    tr.Outcome = Outcome.Failed;
                    break;
                case Microsoft.VisualStudio.TestPlatform.ObjectModel.TestOutcome.Skipped:
                    tr.Outcome = Outcome.Skipped;
                    break;
                case Microsoft.VisualStudio.TestPlatform.ObjectModel.TestOutcome.NotFound:
                    tr.Outcome = Outcome.NotFound;
                    break;
                default:
                    tr.Outcome = Outcome.None;
                    break;
            }

            if (test.Messages != null && test.Messages.Count > 0)
            {
                foreach (var message in test.Messages)
                {
                    tr.TestResultMessages.Add(new TestResultMessage()
                    {
                        Category = message.Category,
                        Text = message.Text,
                    });
                }
            }

            testRun.TestResults.Add(tr);
            SaveChanges();
        }

        /// <summary>
        /// VSTestLogger method being called when testing is done
        /// </summary>
        /// <param name="stats">Test statistocs</param>
        /// <param name="isCanceled">Was the test cancelled</param>
        /// <param name="isAborted">Was the test aborted</param>
        /// <param name="error">Was there an error</param>
        /// <param name="attachmentSets">Test attachements</param>
        /// <param name="elapsedTime">How long did the test run</param>
        public void TestAreDone(ITestRunStatistics stats, bool isCanceled, bool isAborted, Exception error, Collection<Microsoft.VisualStudio.TestPlatform.ObjectModel.AttachmentSet> attachmentSets, TimeSpan elapsedTime)
        {
            testRun.TestWasCancelled = isCanceled;
            testRun.TestWasAborted = isAborted;
            testRun.TestTime = elapsedTime;
            testRun.Status = RunStatus.Done;

            // count tests and store summary in the TestRun row
            var passedTests = testRun.TestResults.Where(r => r.Outcome == Outcome.Passed).ToList().Count();
            var failedTests = testRun.TestResults.Where(r => r.Outcome == Outcome.Failed).ToList().Count();
            var skippedTests = testRun.TestResults.Where(r => r.Outcome == Outcome.Skipped).ToList().Count();
            var notFoundTests = testRun.TestResults.Where(r => r.Outcome == Outcome.NotFound).ToList().Count();

            testRun.TestsPassed = passedTests;
            testRun.TestsFailed = failedTests;
            testRun.TestsSkipped = skippedTests;
            testRun.TestsNotFound = notFoundTests;

            SaveChanges();
        }
        #endregion

        #region private methods
        /// <summary>
        /// Grabs the build number of the environment that we're testing
        /// </summary>
        /// <returns>Build number of the environment that's being tested</returns>
        private string GetBuildNumber()
        {
            string build;
            try
            {
                AuthenticationManager am = new AuthenticationManager();
                if (testConfiguration.TestAuthentication.AppOnly)
                {
                    string realm = TokenHelper.GetRealmFromTargetUrl(new Uri(testConfiguration.TestSiteUrl));

                    using (ClientContext ctx = am.GetAppOnlyAuthenticatedContext(testConfiguration.TestSiteUrl, realm, testConfiguration.TestAuthentication.AppId, testConfiguration.TestAuthentication.AppSecret))
                    {
                        ctx.Load(ctx.Web, w => w.Title);
                        ctx.ExecuteQueryRetry();
                        build = ctx.ServerLibraryVersion.ToString();
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(testConfiguration.TestAuthentication.CredentialManagerLabel))
                    {
                        var credentials = CredentialManager.GetSharePointOnlineCredential(testConfiguration.TestAuthentication.CredentialManagerLabel);
                        using (ClientContext ctx = new ClientContext(testConfiguration.TestSiteUrl))
                        {
                            ctx.Credentials = credentials;
                            ctx.Load(ctx.Web, w => w.Title);
                            ctx.ExecuteQueryRetry();
                            build = ctx.ServerLibraryVersion.ToString();
                        }
                    }
                    else
                    {
                        if (testConfiguration.TestAuthentication.Type == TestAuthenticationType.Online)
                        {
                            using (ClientContext ctx = am.GetSharePointOnlineAuthenticatedContextTenant(testConfiguration.TestSiteUrl, testConfiguration.TestAuthentication.User, testConfiguration.TestAuthentication.Password))
                            {
                                ctx.Load(ctx.Web, w => w.Title);
                                ctx.ExecuteQueryRetry();
                                build = ctx.ServerLibraryVersion.ToString();
                            }
                        }
                        else
                        {
                            using (ClientContext ctx = am.GetNetworkCredentialAuthenticatedContext(testConfiguration.TestSiteUrl, testConfiguration.TestAuthentication.User, testConfiguration.TestAuthentication.Password, testConfiguration.TestAuthentication.Domain))
                            {
                                ctx.Load(ctx.Web, w => w.Title);
                                ctx.ExecuteQueryRetry();
                                build = ctx.ServerLibraryVersion.ToString();
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR: Most likely something is wrong with the provided credentials (username+pwd, appid+secret, credential manager setting) causing the below error:");
                Console.WriteLine(ex.ToString());
                throw;
            }

            return build;
        }

        /// <summary>
        /// Persists changes using the entity framework. Puts detailed DbEntityValidationException errors in console
        /// </summary>
        private void SaveChanges()
        {
            try
            {
                context.SaveChanges();
            }
            catch (DbEntityValidationException e)
            {
                foreach (var eve in e.EntityValidationErrors)
                {
                    Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                        eve.Entry.Entity.GetType().Name, eve.Entry.State);
                    foreach (var ve in eve.ValidationErrors)
                    {
                        Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"",
                            ve.PropertyName, ve.ErrorMessage);
                    }
                }
                throw;
            }
        }

        /// <summary>
        /// Grabs a parameter from the parameter collection
        /// </summary>
        /// <param name="parameter">Name of the parameter to fetch</param>
        /// <returns>Value of the parameter</returns>
        private string GetParameter(string parameter)
        {
            if (loggerParameters.ContainsKey(parameter))
            {
                return loggerParameters[parameter];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Decodes a base64 encoded string
        /// </summary>
        /// <param name="base64EncodedData">base64 encoded string with special "=" being replaced by "&equal;"</param>
        /// <returns>Decoded string</returns>
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData.Replace("&equal", "="));
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        #endregion

    }
}
