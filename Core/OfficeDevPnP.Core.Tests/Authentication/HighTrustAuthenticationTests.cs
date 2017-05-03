using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeDevPnP.Core.Tests.Authentication
{
#if ONPREMISES
    [TestClass]
    public class HighTrustAuthenticationTests
    {
        #region Test initialization
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
        }
        #endregion

        [TestMethod]
        public void HighTrustCertificateAppOnlyAuthenticationTest()
        {
            string siteUrl = TestCommon.DevSiteUrl;
            string clientId = TestCommon.AppId;
            string certificatePath = TestCommon.HighTrustCertificatePath;
            string certificatePassword = TestCommon.HighTrustCertificatePassword;
            string certificateIssuerId = TestCommon.HighTrustIssuerId;

            if (String.IsNullOrEmpty(clientId) ||
                String.IsNullOrEmpty(certificatePath) ||
                String.IsNullOrEmpty(certificatePassword) ||
                String.IsNullOrEmpty(certificateIssuerId) ||
                String.IsNullOrEmpty(siteUrl))
            {
                Assert.Inconclusive("Not enough information to execute this test is passed via the app.config file.");
            }

            ClientContext cc = null;

            try
            {
                // Instantiate a ClientContext object based on the defined high trust certificate
                cc = new AuthenticationManager().GetHighTrustCertificateAppOnlyAuthenticatedContext(siteUrl, clientId, certificatePath, certificatePassword, certificateIssuerId);

                // Check if we can read a property from the site
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQueryRetry();
                Console.WriteLine(String.Format("Site title: {0}", cc.Web.Title));

                Assert.IsFalse(string.IsNullOrEmpty(cc.Web.Title), "Unable to retrieve site title");
                // Nothing blew up...so we're good :-)
            }
            finally
            {
                cc.Dispose();
            }
        }
    }
#endif
}
