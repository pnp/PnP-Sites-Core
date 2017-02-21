using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Providers.Extensibility;
using System.Security.Cryptography.X509Certificates;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using Microsoft.SharePoint.Client;
using System.Threading;

namespace OfficeDevPnP.Core.Tests.Framework.Providers
{
    [TestClass]
    public class XMLProvidersTests
    {
        #region Test variables

        static string testContainer = "pnptest";
        static string testContainerSecure = "pnptestsecure";
        static string testTemplatesDocLib = "PnPTemplatesTests";

        private const string TEST_CATEGORY = "Framework Provisioning XML Providers";

        #endregion

        #region Test initialize and cleanup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            if (!String.IsNullOrEmpty(TestCommon.AzureStorageKey))
            {
                UploadTemplatesToAzureStorageAccount();
            }

            if (!String.IsNullOrEmpty(TestCommon.DevSiteUrl))
            {
                CleanupTemplatesFromSharePointLibrary();
                UploadTemplatesToSharePointLibrary();
            }
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            if (!String.IsNullOrEmpty(TestCommon.AzureStorageKey))
            {
                CleanupTemplatesFromAzureStorageAccount();
            }

            if (!String.IsNullOrEmpty(TestCommon.DevSiteUrl))
            {
                CleanupTemplatesFromSharePointLibrary();
            }
        }

        private static void CleanupTemplatesFromAzureStorageAccount()
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(TestCommon.AzureStorageKey);
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            CloudBlobContainer container = blobClient.GetContainerReference(testContainer);
            container.DeleteIfExists();

            CloudBlobContainer containerSecure = blobClient.GetContainerReference(testContainerSecure);
            containerSecure.DeleteIfExists();
        }

        private static void UploadTemplatesToAzureStorageAccount()
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(TestCommon.AzureStorageKey);
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            CloudBlobContainer container = blobClient.GetContainerReference(testContainer);
            // Create the container if it doesn't already exist.
            container.CreateIfNotExists();

            // Upload files
            CloudBlockBlob blockBlob = container.GetBlockBlobReference("ProvisioningTemplate-2015-03-Sample-01.xml");
            // Create or overwrite the "myblob" blob with contents from a local file.
            using (var fileStream = System.IO.File.OpenRead(String.Format(@"{0}\..\..\Resources\Templates\{1}", AppDomain.CurrentDomain.BaseDirectory, "ProvisioningTemplate-2015-03-Sample-01.xml")))
            {
                blockBlob.UploadFromStream(fileStream);
            }

            blockBlob = container.GetBlockBlobReference("ProvisioningTemplate-2015-03-Sample-02.xml");
            // Create or overwrite the "myblob" blob with contents from a local file.
            using (var fileStream = System.IO.File.OpenRead(String.Format(@"{0}\..\..\Resources\Templates\{1}", AppDomain.CurrentDomain.BaseDirectory, "ProvisioningTemplate-2015-03-Sample-02.xml")))
            {
                blockBlob.UploadFromStream(fileStream);
            }

            CloudBlobContainer containerSecure = blobClient.GetContainerReference(testContainerSecure);
            // Create the container if it doesn't already exist.
            containerSecure.CreateIfNotExists();

            // Avoid public access to this test container
            BlobContainerPermissions bcp = new BlobContainerPermissions();
            bcp.PublicAccess = BlobContainerPublicAccessType.Off;
            containerSecure.SetPermissions(bcp);

            blockBlob = containerSecure.GetBlockBlobReference("ProvisioningTemplate-2015-03-Sample-02.xml");
            // Create or overwrite the "myblob" blob with contents from a local file.
            using (var fileStream = System.IO.File.OpenRead(String.Format(@"{0}\..\..\Resources\Templates\{1}", AppDomain.CurrentDomain.BaseDirectory, "ProvisioningTemplate-2015-03-Sample-02.xml")))
            {
                blockBlob.UploadFromStream(fileStream);
            }
        }

        private static void UploadTemplatesToSharePointLibrary()
        {
            var context = TestCommon.CreatePnPClientContext();

            var docLib = context.Web.CreateDocumentLibrary(testTemplatesDocLib);
            context.Load(docLib, d => d.RootFolder);
            context.ExecuteQueryRetry();

            var templatesToUpload = new string[] {
                "ProvisioningTemplate-2016-05-Sample-02.xml"
            };

            foreach (var tempFile in templatesToUpload)
            {
                // Create or overwrite the "myblob" blob with contents from a local file.
                using (var fileStream = System.IO.File.OpenRead(String.Format(@"{0}\..\..\Resources\Templates\{1}", AppDomain.CurrentDomain.BaseDirectory, tempFile)))
                {
                    docLib.RootFolder.UploadFile(tempFile, fileStream, true);
                    context.ExecuteQueryRetry();
                }
            }

        }

        private static void CleanupTemplatesFromSharePointLibrary()
        {
            var context = TestCommon.CreatePnPClientContext();

            var docLib = context.Web.GetListByTitle(testTemplatesDocLib);
            if (docLib != null)
            {
                context.Load(docLib);
                context.ExecuteQueryRetry();
                docLib.DeleteObject();
                context.ExecuteQueryRetry();
            }
        }
        #endregion

        #region XML File System tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplatesTest()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplates();

            Assert.IsTrue(result.Count == 10);
            Assert.IsTrue(result[0].Files.Count == 5);
            Assert.IsTrue(result[1].Files.Count == 5);
            Assert.IsTrue(result[2].Files.Count == 6);
            Assert.IsTrue(result[3].Files.Count == 5);
            Assert.IsTrue(result[4].Files.Count == 1);
            Assert.IsTrue(result[5].Files.Count == 5);
            Assert.IsTrue(result[6].Files.Count == 1);
            Assert.IsTrue(result[7].Files.Count == 1);
            Assert.IsTrue(result[8].Files.Count == 1);
            Assert.IsTrue(result[9].Files.Count == 5);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplate1Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-01.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 1);
            Assert.IsTrue(result.Files.Count == 1);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplate2Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.Files.Count == 5);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplate3Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        #endregion

        #region XML SharePoint tests
          
        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSharePointGetTemplate1Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            var context = TestCommon.CreatePnPClientContext();

            XMLSharePointTemplateProvider provider =
                new XMLSharePointTemplateProvider(context,
                    TestCommon.DevSiteUrl,
                    testTemplatesDocLib);

            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        #endregion

        #region XML Azure Storage tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLAzureStorageGetTemplatesTest()
        {
            if (String.IsNullOrEmpty(TestCommon.AzureStorageKey))
            {
                Assert.Inconclusive("No Azure Storage Key defined in App.Config, so can't test");
            }

            XMLTemplateProvider provider =
                new XMLAzureStorageTemplateProvider(
                    TestCommon.AzureStorageKey, testContainer);

            var result = provider.GetTemplates();

            Assert.IsTrue(result.Count == 2);
            Assert.IsTrue(result[0].Files.Count == 1);
            Assert.IsTrue(result[1].Files.Count == 5);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLAzureStorageGetTemplate1Test()
        {
            if (String.IsNullOrEmpty(TestCommon.AzureStorageKey))
            {
                Assert.Inconclusive("No Azure Storage Key defined in App.Config, so can't test");
            }

            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLAzureStorageTemplateProvider(
                    TestCommon.AzureStorageKey, testContainer);

            var result = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-01.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 1);
            Assert.IsTrue(result.Files.Count == 1);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLAzureStorageGetTemplate2SecureTest()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            if (String.IsNullOrEmpty(TestCommon.AzureStorageKey))
            {
                Assert.Inconclusive("No Azure Storage Key defined in App.Config, so can't test");
            }

            XMLTemplateProvider provider =
                new XMLAzureStorageTemplateProvider(
                    TestCommon.AzureStorageKey, testContainerSecure);

            var result = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.Files.Count == 5);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemConvertTemplatesFromV201503toV201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var _source1 = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-01.xml");
            provider.SaveAs(_source1, "ProvisioningTemplate-2016-05-Sample-01.xml", XMLPnPSchemaFormatter.LatestFormatter);

            var _source2 = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-02.xml");
            provider.SaveAs(_source2, "ProvisioningTemplate-2016-05-Sample-02.xml", XMLPnPSchemaFormatter.LatestFormatter);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void ResolveSchemaFormatV201503()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.Files.Count == 5);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void ResolveSchemaFormatV201505()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningSchema-2015-05-ReferenceSample-01.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.Files.Count == 5);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        #endregion

        #region XInclude Tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLResolveValidXInclude()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2015-05-Valid-XInclude-01.xml");

            Assert.IsTrue(result.PropertyBagEntries.Count == 2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLResolveInvalidXInclude()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2015-05-NOT-Valid-XInclude-01.xml");

            Assert.IsTrue(result.PropertyBagEntries.Count == 0);
        }

        #endregion

        #region Provider Extensibility Tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLEncryptionTest()
        {
            X509Certificate2 certificate = RetrieveCertificateFromStore(new X509Store(StoreLocation.CurrentUser), "PnPTestCertificate");

            if (certificate == null)
            {
                Assert.Inconclusive("Missing certificate with SN=PnPTestCertificate in CurrentUser Certificate Store, so can't test");
            }

            XMLEncryptionTemplateProviderExtension extension =
                new XMLEncryptionTemplateProviderExtension();

            extension.Initialize(certificate);

            ITemplateProviderExtension[] extensions = new ITemplateProviderExtension[] { extension };

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-01.xml");
            template.DisplayName = "Ciphered template";

            provider.SaveAs(template, "ProvisioningTemplate-2016-05-Ciphered.xml", extensions);
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Ciphered.xml", extensions);

            provider.Delete("ProvisioningTemplate-2016-05-Ciphered.xml");

            Assert.IsTrue(result.DisplayName == "Ciphered template");
        }

        private static X509Certificate2 RetrieveCertificateFromStore(X509Store store, String subjectName)
        {
            if (store == null)
                throw new ArgumentNullException("store");

            X509Certificate2 cert = null;

            try
            {
                store.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certs = store.Certificates.Find(X509FindType.FindBySubjectName, subjectName, false);

                if (certs.Count > 0)
                {
                    // Get the first certificate in the collection
                    cert = certs[0];
                }
            }
            finally
            {
                if (store != null)
                    store.Close();
            }

            return cert;
        }

        #endregion
    }
}
