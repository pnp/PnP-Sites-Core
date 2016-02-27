using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Connectors
{
    [TestClass]
    public class ConnectorZipFileTests
    {
        #region Test initialize and cleanup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
			// File system setup
			if (System.IO.File.Exists(@".\Resources\Templates\sitetemplate-new.zip"))
			{
				System.IO.File.Delete(@".\Resources\Templates\sitetemplate-new.zip");
            }

			if (System.IO.File.Exists(@".\Resources\Templates\sitetemplate.zip"))
			{
				ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "");
				if (connector.GetFileStream("blabla.png") != null)
				{
					connector.DeleteFile("blabla.png");
				}
				if (connector.GetFileStream("blabla.png", "newfolder") != null)
				{
					connector.DeleteFile("blabla.png", "newfolder");
				}
			}
		}
        #endregion

        #region File connector tests
        /// <summary>
        /// Get file as string from provided directory and folder. Specify both directory and container
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorGetFile1Test()
        {
            ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "ProvisioningSchema-2015-12-FullSample-02");

            string file = connector.GetFile("ProvisioningSchema-2015-12-FullSample-02.xml");
            Assert.IsNotNull(file);

            string file2 = connector.GetFile("Idonotexist.zip");
            Assert.IsNull(file2);
        }

		/// <summary>
		/// Get file as string from provided directory and folder. Specify only directory and container, but override the container in the GetFile method
		/// </summary>
		[TestMethod]
        public void ZipFileConnectorGetFile2Test()
        {
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "wrong");

            string file = connector.GetFile("ProvisioningSchema-2015-05-FullSample-01.xml", "ProvisioningSchema-2015-05-FullSample-01");
            Assert.IsNotNull(file);

            string file2 = connector.GetFile("ProvisioningSchema-2015-05-FullSample-01.xml");
            Assert.IsNull(file2);
        }

        /// <summary>
        /// Get files in the specified directory
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorGetFiles1Test()
        {
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "ProvisioningSchema-2015-05-FullSample-01");
            var files = connector.GetFiles();
            Assert.IsTrue(files.Count > 0);
        }

        /// <summary>
        /// Get files in the specified directory, override the set container in the GetFiles method
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorGetFiles2Test()
        {
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "");
            var files = connector.GetFiles("ProvisioningSchema-2015-05-FullSample-01");
            Assert.IsTrue(files.Count > 0);

            var files2 = connector.GetFiles("ProvisioningTemplate-2015-03-Samples");
            Assert.IsTrue(files2.Count > 0);
        }

        /// <summary>
        /// Get file as stream.
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorGetFileBytes1Test()
        {
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "");

            using (var bytes = connector.GetFileStream("ProvisioningSchema-2015-05-ReferenceSample-01.json", "ProvisioningSchema-2015-05-FullSample-01"))
            {
                Assert.IsTrue(bytes.Length > 0);
            }

            using (var bytes2 = connector.GetFileStream("Idonotexist.xml"))
            {
                Assert.IsNull(bytes2);
            }
        }

        /// <summary>
        /// Save file to default container
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorSaveStream1Test()
        {
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "");
            long byteCount = 0;
            using (var fileStream = System.IO.File.OpenRead(@".\resources\office365.png"))
            {
                byteCount = fileStream.Length;
				connector.SaveFileStream("blabla.png", fileStream);
            }

            //read the file
            using (var bytes = connector.GetFileStream("blabla.png"))
            {
                Assert.IsTrue(byteCount == bytes.Length);
            }
            // file will be deleted at end of test 
        }

        /// <summary>
        /// Save file to specified container using a non existing folder...folder will be created on the fly
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorSaveStream2Test()
        {
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "");
            long byteCount = 0;
            using (var fileStream = System.IO.File.OpenRead(@".\resources\office365.png"))
            {
                byteCount = fileStream.Length;
				connector.SaveFileStream("blabla.png", "newfolder", fileStream);
            }

            //read the file
            using (var bytes = connector.GetFileStream("blabla.png", "newfolder"))
            {
                Assert.IsTrue(byteCount == bytes.Length);
            }

            // file will be deleted at end of test 
        }

        /// <summary>
        /// Save file to specified container, check if overwrite works
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorSaveStream3Test()
        {
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "wrong");
            using (var fileStream = System.IO.File.OpenRead(@".\resources\custombg.jpg"))
            {
				connector.SaveFileStream("blabla.png", "", fileStream);
            }

            long byteCount = 0;
            using (var fileStream = System.IO.File.OpenRead(@".\resources\office365.png"))
            {
                byteCount = fileStream.Length;
				connector.SaveFileStream("blabla.png", "", fileStream);
            }

            //read the file
            using (var bytes = connector.GetFileStream("blabla.png", ""))
            {
                Assert.IsTrue(byteCount == bytes.Length);
            }

            // file will be deleted at end of test 
        }

		/// <summary>
		/// Save file to new zip archive container, check if archive is created if not exists
		/// </summary>
		[TestMethod]
		public void ZipFileConnectorSaveStream4Test()
		{
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate-new.zip", "");
			long byteCount = 0;
			using (var fileStream = System.IO.File.OpenRead(@".\resources\office365.png"))
			{
				byteCount = fileStream.Length;
				connector.SaveFileStream("blabla.png", fileStream);
			}

			//read the file
			using (var bytes = connector.GetFileStream("blabla.png"))
			{
				Assert.IsTrue(byteCount == bytes.Length);
			}

		}

		/// <summary>
		/// Save file to default container
		/// </summary>
		[TestMethod]
        public void ZipFileConnectorDelete1Test()
        {
			// upload the file
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "");
            using (var fileStream = System.IO.File.OpenRead(@".\resources\office365.png"))
            {
				connector.SaveFileStream("blabla.png", fileStream);
            }

			// delete the file
			connector.DeleteFile("blabla.png");

            //read the file
            using (var bytes = connector.GetFileStream("blabla.png"))
            {
                Assert.IsNull(bytes);
            }

            // file will be deleted at end of test 
        }

        /// <summary>
        /// Save file to default container
        /// </summary>
        [TestMethod]
        public void ZipFileConnectorDelete2Test()
        {
			// upload the file
			ZipFileConnector connector = new ZipFileConnector(@".\Resources\Templates\sitetemplate.zip", "wrong");
            using (var fileStream = System.IO.File.OpenRead(@".\resources\office365.png"))
            {
				connector.SaveFileStream("blabla.png", "", fileStream);
            }

			// delete the file
			connector.DeleteFile("blabla.png", "");

            //read the file
            using (var bytes = connector.GetFileStream("blabla.png", ""))
            {
                Assert.IsNull(bytes);
            }

            // file will be deleted at end of test 
        }
        #endregion
    }
}
