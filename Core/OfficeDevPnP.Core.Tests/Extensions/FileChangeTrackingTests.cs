#if !NETSTANDARD2_0
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Utilities;
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
    // Dropped this functionality
    // Used tables 
    /*
     USE[PnP]
    GO

    SET ANSI_NULLS ON
    GO

    SET QUOTED_IDENTIFIER ON
    GO

    CREATE TABLE[dbo].[FileTrackingBaselineSet]
        (

       [Id][int] IDENTITY(1,1) NOT NULL,

      [FileName] [nvarchar] (max) NOT NULL,
	    [Build] [nvarchar] (max) NOT NULL,
	    [FileHash] [nvarchar] (max) NOT NULL,
	    [ChangeDate]
        [datetime]
        NOT NULL,

        [FileContents] [varbinary] (max) NOT NULL,
     CONSTRAINT[PK_FileTrackingBaselineSet] PRIMARY KEY CLUSTERED
    (
       [Id] ASC
    )WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
    )
    GO

 USE [PnP]
GO

    SET ANSI_NULLS ON
    GO

    SET QUOTED_IDENTIFIER ON
    GO

    CREATE TABLE[dbo].[FileTrackingSet]
        (

       [Id][int] IDENTITY(1,1) NOT NULL,

      [TestDate] [datetime]
        NOT NULL,

      [Build] [nvarchar] (max) NOT NULL,
	    [FileName] [nvarchar] (max) NOT NULL,
	    [FileHash] [nvarchar] (max) NOT NULL,
	    [FileChanged]
        [bit]
        NOT NULL,

        [TestSiteUrl] [nvarchar] (max) NOT NULL,
	    [TestUser]
        [nvarchar]
        (max) NULL,

        [TestAppId] [nvarchar]
        (max) NULL,

        [TestComputerName] [nvarchar]
        (max) NULL,
     CONSTRAINT[PK_FileTrackingSet] PRIMARY KEY CLUSTERED
    (
       [Id] ASC
    )WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
    )
    GO




     */

    //[TestClass()]
    //public class FileChangeTrackingTests
    //{
    //    #region Test initialize and cleanup
    //    [TestInitialize]
    //    public void Initialize()
    //    {
    //        if (!TestCommon.TestAutomationSQLDatabaseAvailable())
    //        {
    //            Assert.Inconclusive("No test automation SQL database information found...or found database is not reachable.");
    //        }
    //    }
    //    #endregion


    //    [TestMethod]
    //    public void OOBMasterPagesHaveChangedTest()
    //    {
    //        using (var context = TestCommon.CreateClientContext())
    //        {
    //            var web = context.Web;
    //            //need to get the server relative url 
    //            context.Load(web, w => w.ServerRelativeUrl);
    //            context.ExecuteQueryRetry();

    //            string masterpageGalleryServerRelativeUrl = UrlUtility.Combine(UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl), "_catalogs/masterpage/");
    //            // Test seattle.master
    //            TestFile(context, UrlUtility.Combine(masterpageGalleryServerRelativeUrl, "seattle.master"));
    //            // Test oslo.master
    //            TestFile(context, UrlUtility.Combine(masterpageGalleryServerRelativeUrl, "oslo.master"));

    //        }
    //    }


    //    private void TestFile(ClientContext ctx, string serverRelativeFileUrl)
    //    {
    //        // grab file reference
    //        var file = ctx.Web.GetFileByServerRelativeUrl(serverRelativeFileUrl);
    //        ctx.Load(file);
    //        ctx.ExecuteQueryRetry();

    //        // download file
    //        ClientResult<Stream> data = file.OpenBinaryStream();
    //        ctx.Load(file);
    //        ctx.ExecuteQueryRetry();

    //        // copy to MemoryStream
    //        using (MemoryStream memStream = new MemoryStream())
    //        {
    //            data.Value.CopyTo(memStream);

    //            // compute a hash of the file 
    //            var hashAlgorithm = HashAlgorithm.Create();
    //            // Copy bytes to byte array, getting hash directy from memorystream did not work properly!
    //            byte[] bytes = memStream.ToArray();
    //            byte[] hash = hashAlgorithm.ComputeHash(bytes);
    //            // convert to a hex string
    //            string hex = BitConverter.ToString(hash);

    //            using (SqlConnection connection = new SqlConnection(TestCommon.TestAutomationDatabaseConnectionString))
    //            {
    //                string appId = ConfigurationManager.AppSettings["AppId"];
    //                string user = null;
    //                if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOCredentialManagerLabel"]))
    //                {
    //                    user = CredentialManager.GetCredential(ConfigurationManager.AppSettings["SPOCredentialManagerLabel"]).UserName;
    //                }
    //                else
    //                {
    //                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOUserName"]))
    //                    {
    //                        user = ConfigurationManager.AppSettings["SPOUserName"];
    //                    }
    //                    else if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremUserName"]) && !String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremDomain"]))
    //                    {
    //                        user = string.Format("{0}\\{1}", ConfigurationManager.AppSettings["OnPremDomain"], ConfigurationManager.AppSettings["OnPremUserName"]);
    //                    }
    //                }


    //                string fileHashToCompareAgainst = "";
    //                using (SqlCommand command = new SqlCommand("SELECT TOP 1 [Id], [FileName], [Build], [FileHash], [ChangeDate] FROM [dbo].[FileTrackingBaselineSet] WHERE [FileName] = @FileName AND [Build] = @Build ORDER BY [ChangeDate] DESC", connection))
    //                {
    //                    command.Parameters.AddWithValue("@FileName", serverRelativeFileUrl);
    //                    command.Parameters.AddWithValue("@Build", ctx.ServerLibraryVersion.ToString());

    //                    connection.Open();
    //                    using (SqlDataReader reader = command.ExecuteReader())
    //                    {
    //                        while (reader.Read())
    //                        {
    //                            fileHashToCompareAgainst = reader["FileHash"].ToString();
    //                        }
    //                    }
    //                }

    //                DateTime nowDate = DateTime.Now;

    //                // if there's no baseline record yet then add it
    //                if (String.IsNullOrEmpty(fileHashToCompareAgainst))
    //                {
    //                    using (SqlCommand command = new SqlCommand("INSERT INTO [dbo].[FileTrackingBaselineSet] ([FileName], [Build], [FileHash], [ChangeDate], [FileContents]) VALUES (@FileName, @Build, @FileHash, @ChangeDate, @FileContents)", connection))
    //                    {
    //                        command.Parameters.AddWithValue("@ChangeDate", nowDate);
    //                        command.Parameters.AddWithValue("@Build", ctx.ServerLibraryVersion.ToString());
    //                        command.Parameters.AddWithValue("@FileName", serverRelativeFileUrl);
    //                        command.Parameters.AddWithValue("@FileHash", hex);
    //                        command.Parameters.AddWithValue("@FileContents", memStream.ToArray());

    //                        // insert record
    //                        command.ExecuteNonQuery();
    //                    }
    //                }
    //                else
    //                {
    //                    // add a new comparison record when there was a change detected
    //                    if (!hex.Equals(fileHashToCompareAgainst, StringComparison.InvariantCultureIgnoreCase))
    //                    {
    //                        using (SqlCommand command = new SqlCommand("INSERT INTO [dbo].[FileTrackingBaselineSet] ([FileName], [Build], [FileHash], [ChangeDate], [FileContents]) VALUES (@FileName, @Build, @FileHash, @ChangeDate, @FileContents)", connection))
    //                        {
    //                            command.Parameters.AddWithValue("@ChangeDate", nowDate);
    //                            command.Parameters.AddWithValue("@Build", ctx.ServerLibraryVersion.ToString());
    //                            command.Parameters.AddWithValue("@FileName", serverRelativeFileUrl);
    //                            command.Parameters.AddWithValue("@FileHash", hex);
    //                            command.Parameters.AddWithValue("@FileContents", memStream.ToArray());

    //                            // insert record
    //                            command.ExecuteNonQuery();
    //                        }
    //                    }
    //                }

    //                // prep insert command
    //                using (SqlCommand command = new SqlCommand("INSERT INTO [dbo].[FileTrackingSet] VALUES (@TestDate, @Build, @FileName, @FileHash, @FileChanged, @TestSiteUrl, @TestUser, @TestAppId, @TestComputerName)", connection))
    //                {
    //                    bool hasChanged = false;

    //                    if (!String.IsNullOrEmpty(fileHashToCompareAgainst))
    //                    {
    //                        hasChanged = !hex.Equals(fileHashToCompareAgainst, StringComparison.InvariantCultureIgnoreCase);
    //                    }

    //                    command.Parameters.AddWithValue("@TestDate", nowDate);
    //                    command.Parameters.AddWithValue("@Build", ctx.ServerLibraryVersion.ToString());
    //                    command.Parameters.AddWithValue("@FileName", serverRelativeFileUrl);
    //                    command.Parameters.AddWithValue("@FileHash", hex);
    //                    command.Parameters.AddWithValue("@FileChanged", hasChanged);
    //                    command.Parameters.AddWithValue("@TestSiteUrl", ConfigurationManager.AppSettings["SPODevSiteUrl"]);
    //                    command.Parameters.AddWithValue("@TestUser", user != null ? user : "");
    //                    command.Parameters.AddWithValue("@TestAppId", appId != null ? appId : "");
    //                    command.Parameters.AddWithValue("@TestComputerName", Environment.MachineName);

    //                    // insert record
    //                    command.ExecuteNonQuery();
    //                    connection.Close();
    //                }
    //            }
    //        }
    //    }

    //}
}
#endif