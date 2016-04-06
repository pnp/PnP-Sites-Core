using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Configuration;

namespace OfficeDevPnP.Core.Tests.Framework.TimerJobs
{
    [TestClass]
    public class TimerJobTests
    {
        [TestMethod]
        public void TestMethod1()
        {
            var job = new TestTimerJob("Test");

            job.AddSite(ConfigurationManager.AppSettings["SPODevSiteUrl"]);
            job.UseAzureADAppOnlyAuthentication("13f8eaae-480a-4cb7-b3d2-efed51c640a3", "erwinmcm.com", "c:\\temp\\pnppartnerpackcertificate.pfx", Core.Utilities.EncryptionUtility.ToSecureString("sinterKlaas18"));
           // job.UseOffice365Authentication("erwinmcm");
            job.Run();
        }
    }
}
