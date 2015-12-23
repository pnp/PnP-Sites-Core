using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Operations;
using OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.SQL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildTest
{
    class Program
    {
        static void Main(string[] args)
        {

            AppManager am = new AppManager("https://bertonline.sharepoint.com/sites/dev", AuthenticationType.Office365, "bertonline");

            #region Deploy provider testing
            //am.DeployProviderHostedAppAsAzureWebSite(@"C:\temp\providerhostedapp1\providerhostedapp1\providerhostedapp1.csproj", @"C:\temp\providerhostedapp1\providerhostedapp1Web\providerhostedapp1Web.csproj",
            //                            "d643a56c-5319-48a8-a97f-d7c8b905dac5", "jbaTpSGkvgmK7+BHoS8io22UP/fVUW/UfvEioHfpaRc=",
            //                            "bjansen-automation1.scm.azurewebsites.net:443", "https://bjansen-automation1.azurewebsites.net", "bjansen-automation1",
            //                            @"C:\temp\providerhostedapp1\bjansen-automation1.publishsettings", @"c:\temp\providerhostedapp1.Package");
            #endregion

            #region App Package generation testing
            // Provider hosted testing
            //string appPackageName = "";
            //am.CreateAppPackageForProviderHostedApp(@"C:\temp\Core.EmbedJavaScript\Core.EmbedJavaScript\Core.EmbedJavaScript.csproj",
            //                                        @"C:\temp\Core.EmbedJavaScript\Core.EmbedJavaScriptWeb\Core.EmbedJavaScriptWeb.csproj",
            //                                        "yesImaclientid", "https://localhost:7776", @"c:\temp\Core.EmbedJavaScript.Package", out appPackageName);

            //am.CreateAppPackageForProviderHostedApp(@"C:\temp\providerhostedapp1\providerhostedapp1\providerhostedapp1.csproj", @"C:\temp\providerhostedapp1\providerhostedapp1Web\providerhostedapp1Web.csproj",
            //                            "d643a56c-5319-48a8-a97f-d7c8b905dac5", "https://bjansen-automation1.azurewebsites.net", @"c:\temp\providerhostedapp1.Package", out appPackageName);

            // Sharepoint hosted testing
            //am.CreateAppPackageForSharePointHostedApp(@"C:\temp\SharePointHostedApp1\SharePointHostedApp1\SharePointHostedApp1.csproj", @"c:\temp\SharePointHostedApp1.Package", out appPackageName);
            #endregion

            #region Test Manager testing
            //PnPAppConfigManager p = new PnPAppConfigManager(@"C:\Users\bjansen\Documents\Visual Studio 2013\Projects\MSBuildTests\PnPBuildExtensions\mastertestconfiguration.xml");
            ////Console.WriteLine(p.GetConfigurationElement("OnPremAppOnly", "PnPbranch"));
            //p.GenerateAppConfig("OnlineCred", @"c:\temp");

            //Dictionary<string, string> parameters = new Dictionary<string, string>();
            //parameters.Add("MDPath", @"c:\temp");
            //parameters.Add("PnPConfigurationToTest", "OnlineCred");
            //parameters.Add("PnPBranch", "dev");
            //parameters.Add("PnPBuildConfiguration", "debug");

            //PnPTestManager t = new PnPTestManager(parameters);

            //// Stuff some fake test data
            //TestCase tc1 = new TestCase("OfficeDevPnP.Core.Utilities.Tests.JsonUtilityTests.DeserializeListTest", new Uri("http://www.bing.com"), @"c:\GitHub\BertPnP\OfficeDevPnP.Core\OfficeDevPnP.Core.Tests\Utilities\JsonUtilityTests.cs");
            //TestResult tr1 = new TestResult(tc1);
            //tr1.Outcome = TestOutcome.Passed;
            //tr1.DisplayName = "DeserializeListTest";
            //tr1.Duration = new TimeSpan(0, 0, 0, 0, 245);
            //t.AddTestResult(tr1);

            //TestCase tc2 = new TestCase("OfficeDevPnP.Core.Utilities.Tests.JsonUtilityTests.DeserializeListIsNotFixedSizeTest", new Uri("http://www.bing.com"), @"c:\GitHub\BertPnP\OfficeDevPnP.Core\OfficeDevPnP.Core.Tests\Utilities\JsonUtilityTests.cs");
            //TestResult tr2 = new TestResult(tc2);
            //tr2.Outcome = TestOutcome.Failed;
            //tr2.DisplayName = "DeserializeListIsNotFixedSizeTest";
            //tr2.Duration = new TimeSpan(0, 0, 0, 1, 749);
            //tr2.ErrorMessage = "this is the fake error";
            //tr2.ErrorStackTrace = "this is the stack trace of the error that happened";
            //t.AddTestResult(tr2);

            //Stats s = new Stats();

            ////t.TestAreDone(s, false, false, null, null, new TimeSpan(0, 1, 22));
            //t.GenerateMDSummaryReport();
            #endregion

            string c1 = @"metadata=res://*/SQL.TestModel.csdl|res://*/SQL.TestModel.ssdl|res://*/SQL.TestModel.msl;provider=System.Data.SqlClient;provider connection string=&quot;data source=(localdb)\MSSQLLocalDB;initial catalog=PnPTestAutomation;integrated security=True;MultipleActiveResultSets=True;App=EntityFramework&quot;";

            Console.WriteLine(GetConnectionString(c1));

        }

        private static string GetConnectionString(string c)
        {
            var c2 = c.Substring(c.IndexOf("&quot;") + 6);
            return c2.Substring(0, c2.IndexOf("&quot;"));
        }

    }
}
