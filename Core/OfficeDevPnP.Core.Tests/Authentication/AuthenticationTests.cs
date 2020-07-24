using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using OfficeDevPnP.Core.Entities;
using System.Linq;

#if !NETSTANDARD2_0
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

        /// <summary>
        /// Important: the Azure AD you're using here needs to be consented first, otherwise you'll get an access denied.
        /// Consenting can be done by taking this URL and replacing the client_id parameter value with yours: https://login.microsoftonline.com/common/oauth2/authorize?state=e82ea723-7112-472c-94d4-6e66c0ca52b6&response_type=code+id_token&scope=openid&nonce=c328d2df-43d1-4e4d-a884-7cfb492beadc&client_id=b77caa50-d9ba-4b30-aad6-a40effa2ecd0&redirect_uri=https:%2f%2flocalhost:44304%2fHome%2f&resource=https:%2f%2fgraph.windows.net%2f&prompt=admin_consent&response_mode=form_post
        /// To debug this catch the returned access token and look http://jwt.calebb.net/ to see if the token contains roles claims
        /// </summary>
        [TestMethod]
        public void AzureADAuthFullControlPermissionTest()
        {
            string siteUrl = TestCommon.DevSiteUrl;
            string spoUserName = AuthenticationTests.UserName;
            string azureADCertPfxPassword = TestCommon.AzureADCertPfxPassword;
            string azureADClientId = TestCommon.AzureADClientId;
            string azureADCertificateFilePath = TestCommon.AzureADCertificateFilePath;
            if (string.IsNullOrEmpty(azureADCertificateFilePath))
            {
                azureADCertificateFilePath = @"resources\PnPAzureAppTest.pfx";
            }

            if (String.IsNullOrEmpty(azureADCertificateFilePath) ||
                String.IsNullOrEmpty(azureADCertPfxPassword) ||
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
                    cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, azureADClientId, domain, azureADCertificateFilePath, azureADCertPfxPassword, AzureEnvironment.PPE);
                }
                else
                {
                    cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, azureADClientId, domain, azureADCertificateFilePath, azureADCertPfxPassword);
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
#endif