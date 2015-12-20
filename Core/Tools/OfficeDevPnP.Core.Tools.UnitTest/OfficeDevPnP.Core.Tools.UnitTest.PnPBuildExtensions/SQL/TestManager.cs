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
        #endregion

        #region Constructor
        public TestManager(Dictionary<string, string> parameters)
        {
            //System.Diagnostics.Debugger.Launch();
            loggerParameters = parameters;
            context = new TestModelContainer(Base64Decode(GetParameter("PnPSQLConnectionString")).Replace("&quot;", "\""));

            // find the used configuration
            string configurationToTest = GetParameter("PnPConfigurationToTest");
            testConfiguration = context.TestConfigurationSet.Where(s => s.Name.Equals(configurationToTest, StringComparison.InvariantCultureIgnoreCase)).First();
            if (testConfiguration == null)
            {
                throw new Exception(String.Format("Test configuration with name {0} was not found", configurationToTest));
            }

            // Grab the build of the environment we're testing
            string build = GetBuildNumber();

            testRun = new TestRun()
            {
                TestDate = DateTime.Now,
                TestTime = new TimeSpan(0, 0, 0),
                Build = build,
                TestWasAborted = false,
                TestWasCancelled = false,
                TestConfiguration = testConfiguration,
            };
            context.TestRunSet.Add(testRun);
            SaveChanges();
        }
        #endregion

        #region public methods
        public void AddTestResult(Microsoft.VisualStudio.TestPlatform.ObjectModel.TestResult test)
        {

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

        public void TestAreDone(ITestRunStatistics stats, bool isCanceled, bool isAborted, Exception error, Collection<Microsoft.VisualStudio.TestPlatform.ObjectModel.AttachmentSet> attachmentSets, TimeSpan elapsedTime)
        {
            testRun.TestWasCancelled = isCanceled;
            testRun.TestWasAborted = isAborted;
            testRun.TestTime = elapsedTime;
            SaveChanges();
        }
        #endregion

        #region private methods
        private string GetBuildNumber()
        {
            string build;
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

            return build;
        }

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

        private string GetParameter(string parameter)
        {
            if (loggerParameters.ContainsKey(parameter))
            {
                return loggerParameters[parameter];
            }
            else
            {
                throw new ArgumentException(String.Format("Requested parameter {0} is not defined", parameter));
            }
        }

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData.Replace("&equal", "="));
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        #endregion

    }
}
