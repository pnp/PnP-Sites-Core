using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.AppModelExtensions
{
    [TestClass()]
    public class FileChangeTrackingTests
    {
        #region Test initialize and cleanup
        [TestInitialize]
        public void Initialize()
        {
            if (!TestCommon.TestAutomationSQLDatabaseAvailable())
            {
                Assert.Inconclusive("No test automation SQL database information found...or found database is not reachable.");
            }
        }
        #endregion


        [TestMethod]
        public void OOBMasterPagesHaveChangedTest()
        {
            using (var context = TestCommon.CreateClientContext())
            {
                var web = context.Web;
                //need to get the server relative url 
                context.Load(web, w => w.ServerRelativeUrl);
                context.ExecuteQueryRetry();

                string masterpageGalleryServerRelativeUrl = UrlUtility.Combine(UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl), "_catalogs/masterpage/");
                // Test seattle.master
                TestFile(context, UrlUtility.Combine(masterpageGalleryServerRelativeUrl, "seattle.master"), "56-44-23-D0-20-60-DC-52-87-26-42-E4-D6-0E-4B-01-6C-8A-D5-6A");
                // Test oslo.master
                TestFile(context, UrlUtility.Combine(masterpageGalleryServerRelativeUrl, "oslo.master"), "42-43-6D-C0-88-C6-13-43-2B-01-4D-03-DE-AC-7B-23-80-B4-93-10");

            }
        }


        private void TestFile(ClientContext ctx, string serverRelativeFileUrl, string knownHash)
        {
            // grab file reference
            var file = ctx.Web.GetFileByServerRelativeUrl(serverRelativeFileUrl);
            ctx.Load(file);
            ctx.ExecuteQueryRetry();

            // download file
            ClientResult<Stream> data = file.OpenBinaryStream();
            ctx.Load(file);
            ctx.ExecuteQueryRetry();

            // copy to MemoryStream
            using (MemoryStream memStream = new MemoryStream())
            {
                data.Value.CopyTo(memStream);

                // compute a hash of the file 
                var hashAlgorithm = HashAlgorithm.Create();
                // Copy bytes to byte array, getting hash directy from memorystream did not work properly!
                byte[] bytes = memStream.ToArray();
                byte[] hash = hashAlgorithm.ComputeHash(bytes);
                // convert to a hex string
                string hex = BitConverter.ToString(hash);

                using (SqlConnection connection = new SqlConnection(TestCommon.TestAutomationDatabaseConnectionString))
                {
                    string appId = ConfigurationManager.AppSettings["AppId"];
                    string user = ConfigurationManager.AppSettings["SPOUserName"];

                    // prep insert command
                    using (SqlCommand command = new SqlCommand("INSERT INTO [dbo].[FileTrackingSet] VALUES (@TestDate, @Build, @FileName, @FileHash, @FileChanged, @TestSiteUrl, @TestUser, @TestAppId, @TestComputerName)", connection))
                    {
                        command.Parameters.AddWithValue("@TestDate", DateTime.Now);
                        command.Parameters.AddWithValue("@Build", ctx.ServerLibraryVersion.ToString());
                        command.Parameters.AddWithValue("@FileName", serverRelativeFileUrl);
                        command.Parameters.AddWithValue("@FileHash", hex);
                        //command.Parameters.AddWithValue("@FileBytes", memStream.ToArray());
                        command.Parameters.AddWithValue("@FileChanged", !hex.Equals(knownHash, StringComparison.InvariantCultureIgnoreCase));
                        command.Parameters.AddWithValue("@TestSiteUrl", ConfigurationManager.AppSettings["SPODevSiteUrl"]);
                        command.Parameters.AddWithValue("@TestUser", user != null ? user : "");
                        command.Parameters.AddWithValue("@TestAppId", appId != null ? appId : "");
                        command.Parameters.AddWithValue("@TestComputerName", Environment.MachineName);

                        // insert record
                        connection.Open();
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                }
            }
        }

    }
}
