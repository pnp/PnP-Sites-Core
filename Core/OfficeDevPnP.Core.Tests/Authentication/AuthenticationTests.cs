using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;

namespace OfficeDevPnP.Core.Tests.Authentication
{
#if !ONPREMISES
    [TestClass]
    public class AuthenticationTests
    {
        private static string UserName;


        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                List<UserEntity> admins = clientContext.Web.GetAdministrators();
                UserName = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[2]; 
            }
        }

        [TestMethod]
        public void AzureADAuth1()
        {
            string siteUrl = TestCommon.DevSiteUrl;
            string spoUserName = AuthenticationTests.UserName;
            string azureADCertPfxPassword = TestCommon.AzureADCertPfxPassword;
            string azureADClientId = TestCommon.AzureADClientId;

            if (String.IsNullOrEmpty(azureADCertPfxPassword) ||
                String.IsNullOrEmpty(azureADClientId) ||
                String.IsNullOrEmpty(spoUserName) ||
                String.IsNullOrEmpty(siteUrl))
            {
                Assert.Inconclusive("Not enough information to execute this test is passed via the app.config file.");
            }

            ClientContext cc = null;

            try
            {
                string domain = spoUserName.Split(new string[] { "@" }, StringSplitOptions.RemoveEmptyEntries)[1];

                if (new Uri(siteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, azureADClientId, domain, @"resources\PnPAzureAppTest.pfx", azureADCertPfxPassword, AzureEnvironment.PPE);
                }
                else
                {
                    cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, azureADClientId, domain, @"resources\PnPAzureAppTest.pfx", azureADCertPfxPassword);
                }

                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQueryRetry();
                Console.WriteLine(String.Format("Site title: {0}", cc.Web.Title));
            }
            finally
            {
                cc.Dispose();
            }
        }
    }
#endif
}
