using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Authentication
{
#if !ONPREMISES
    [TestClass]
    public class AuthenticationTests
    {
        private static string UserName;

        #region Test initialization
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                List<UserEntity> admins = clientContext.Web.GetAdministrators();
                UserName = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[2]; 
            }
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                DeleteListsImplementation(clientContext);
            }
        }
        #endregion

        [TestMethod]
        public void AzureADAuthFullControlPermissionTest()
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

                // Instantiate a ClientContext object based on the defined Azure AD application
                if (new Uri(siteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, azureADClientId, domain, @"resources\PnPAzureAppTest.pfx", azureADCertPfxPassword, AzureEnvironment.PPE);
                }
                else
                {
                    cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, azureADClientId, domain, @"resources\PnPAzureAppTest.pfx", azureADCertPfxPassword);
                }

                // Check if we can read a property from the site
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQueryRetry();
                Console.WriteLine(String.Format("Site title: {0}", cc.Web.Title));

                // Verify manage permissions by creating a new list - see https://technet.microsoft.com/en-us/library/cc721640.aspx
                var list = cc.Web.CreateList(ListTemplateType.DocumentLibrary, "Test_list_" + DateTime.Now.ToFileTime(), false);

                // Verify full control by enumerating permissions - see https://technet.microsoft.com/en-us/library/cc721640.aspx
                var roleAssignments = cc.Web.GetAllUniqueRoleAssignments();

                // Nothing blew up...so we're good :-)

            }
            finally
            {
                cc.Dispose();
            }
        }

        #region Helper methods
        private static void DeleteListsImplementation(ClientContext cc)
        {
            cc.Load(cc.Web.Lists, f => f.Include(t => t.Title));
            cc.ExecuteQueryRetry();

            foreach (var list in cc.Web.Lists.ToList())
            {
                if (list.Title.StartsWith("Test_list_"))
                {
                    list.DeleteObject();
                }
            }
            cc.ExecuteQueryRetry();
        }
        #endregion
    }
#endif
}
