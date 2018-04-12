#if !NETSTANDARD2_0
using System;
using System.Linq;
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
using System.Xml.Linq;
using OfficeDevPnP.Core.Utilities;
using System.Collections.Generic;

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

#endregion=

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

            Assert.IsTrue(result.Count > 15);
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
#if !NETSTANDARD2_0
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
#endif

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
            provider.SaveAs(_source1, "ProvisioningTemplate-2016-05-Sample-01.xml", XMLPnPSchemaFormatter.GetSpecificFormatter(XMLPnPSchemaVersion.V201605));

            var _source2 = provider.GetTemplate("ProvisioningTemplate-2015-03-Sample-02.xml");
            provider.SaveAs(_source2, "ProvisioningTemplate-2016-05-Sample-02.xml", XMLPnPSchemaFormatter.GetSpecificFormatter(XMLPnPSchemaVersion.V201605));
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

#region Formatter Refactoring Tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_201605()
        {
            var _expectedID = "SPECIALTEAM-01";
            var _expectedVersion = 1.2;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 1);

            //content types asserts
            Assert.IsNotNull(result.ContentTypes);
            Assert.AreEqual(1, result.ContentTypes.Count);
            Assert.IsNotNull(result.ContentTypes[0].FieldRefs);
            Assert.AreEqual(4, result.ContentTypes[0].FieldRefs.Count);
            Assert.AreEqual(1, result.ContentTypes[0].FieldRefs.Count(f => f.Required));

            Assert.IsTrue(result.PropertyBagEntries.Count == 2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT.xml", serializer);

            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT.xml"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_SerializeDeserialize_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template1 = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template1);

            provider.SaveAs(template1, "ProvisioningTemplate-2016-05-Sample-03-OUT.xml", serializer);
            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT.xml"));

            var template2 = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03-OUT.xml", serializer);
            Assert.IsNotNull(template2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_SerializeDeserialize_201705()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201705Serializer();
            var template1 = provider.GetTemplate("ProvisioningSchema-2017-05-FullSample-01.xml", serializer);
            Assert.IsNotNull(template1);

            provider.SaveAs(template1, "ProvisioningSchema-2017-05-FullSample-01-OUT.xml", serializer);
            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningSchema-2017-05-FullSample-01-OUT.xml"));

            var template2 = provider.GetTemplate("ProvisioningSchema-2017-05-FullSample-01-OUT.xml", serializer);
            Assert.IsNotNull(template2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_SerializeDeserialize_201801()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201801Serializer();
            var template1 = provider.GetTemplate("ProvisioningSchema-2018-01-FullSample-01.xml", serializer);
            Assert.IsNotNull(template1);

            // Add stuff that is not supported anymore, to test the serialization behavior
            template1.AddIns.Add(new AddIn { PackagePath = "test", Source = "test" });

            provider.SaveAs(template1, "ProvisioningSchema-2018-01-FullSample-01-OUT.xml", serializer);
            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningSchema-2018-01-FullSample-01-OUT.xml"));

            var template2 = provider.GetTemplate("ProvisioningSchema-2018-01-FullSample-01-OUT.xml", serializer);
            Assert.IsNotNull(template2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_SerializeDeserialize_201805()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201805Serializer();
            var template1 = provider.GetTemplate("ProvisioningSchema-2018-05-FullSample-01.xml", serializer);
            Assert.IsNotNull(template1);

            // Add stuff that is not supported anymore, to test the serialization behavior
            template1.AddIns.Add(new AddIn { PackagePath = "test", Source = "test" });

            provider.SaveAs(template1, "ProvisioningSchema-2018-05-FullSample-01-OUT.xml", serializer);
            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningSchema-2018-05-FullSample-01-OUT.xml"));

            var template2 = provider.GetTemplate("ProvisioningSchema-2018-05-FullSample-01-OUT.xml", serializer);
            Assert.IsNotNull(template2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ContentTypes_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result.ContentTypes);

            var ct = result.ContentTypes.FirstOrDefault();
            Assert.IsNotNull(ct);
            Assert.AreEqual("0x01005D4F34E4BE7F4B6892AEBE088EDD215E", ct.Id);
            Assert.AreEqual("General Project Document", ct.Name);
            Assert.AreEqual("General Project Document Content Type", ct.Description);
            Assert.AreEqual("Base Foundation Content Types", ct.Group);
            Assert.AreEqual("/Forms/DisplayForm.aspx", ct.DisplayFormUrl);
            Assert.AreEqual("/Forms/NewForm.aspx", ct.NewFormUrl);
            Assert.AreEqual("/Forms/EditForm.aspx", ct.EditFormUrl);
            Assert.AreEqual("DocumentTemplate.dotx", ct.DocumentTemplate);
            Assert.IsTrue(ct.Hidden);
            Assert.IsTrue(ct.Overwrite);
            Assert.IsTrue(ct.ReadOnly);
            Assert.IsTrue(ct.Sealed);

            Assert.IsNotNull(ct.DocumentSetTemplate);
            Assert.IsNotNull(ct.DocumentSetTemplate.AllowedContentTypes);
            Assert.IsNotNull(ct.DocumentSetTemplate.AllowedContentTypes.FirstOrDefault(c => c == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E002"));
            Assert.AreNotEqual(Guid.Empty, ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c == new Guid("f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3")));
            Assert.AreNotEqual(Guid.Empty, ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c == new Guid("a8df65ec-0d06-4df1-8edf-55d48b3936dc")));
            Assert.AreNotEqual(Guid.Empty, ct.DocumentSetTemplate.WelcomePageFields.FirstOrDefault(c => c == new Guid("c69d2ffc-0c86-474a-9cc7-dcd7774da531")));
            Assert.AreNotEqual(Guid.Empty, ct.DocumentSetTemplate.WelcomePageFields.FirstOrDefault(c => c == new Guid("b9132b30-2b9e-47d4-b0fc-1ac34a61506f")));
            Assert.AreEqual("home.aspx", ct.DocumentSetTemplate.WelcomePage);
            Assert.IsNotNull(ct.DocumentSetTemplate.DefaultDocuments);

            var dd = ct.DocumentSetTemplate.DefaultDocuments.FirstOrDefault(d => d.ContentTypeId == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E001");
            Assert.IsNotNull(dd);
            Assert.AreEqual("document.dotx", dd.FileSourcePath);
            Assert.AreEqual("DefaultDocument", dd.Name);

            Assert.IsNotNull(ct.FieldRefs);
            Assert.AreEqual(4, ct.FieldRefs.Count);

            var field = ct.FieldRefs.FirstOrDefault(f => f.Name == "TestField");
            Assert.IsNotNull(field);
            Assert.AreEqual(new Guid("23203e97-3bfe-40cb-afb4-07aa2b86bf45"), field.Id);
            Assert.IsTrue(field.Required);
            Assert.IsTrue(field.Hidden);
        }


        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ContentTypes_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result =  new ProvisioningTemplate();

            var nct = new Core.Framework.Provisioning.Model.ContentType()
            {
                Id = "0x01005D4F34E4BE7F4B6892AEBE088EDD215E",
                Name = "General Project Document",
                Description = "General Project Document Content Type",
                Group = "Base Foundation Content Types",
                DisplayFormUrl = "/Forms/DisplayForm.aspx",
                NewFormUrl = "/Forms/NewForm.aspx",
                EditFormUrl = "/Forms/EditForm.aspx",
                DocumentTemplate = "DocumentTemplate.dotx",
                Hidden = true,
                Overwrite = true,
                ReadOnly = true,
                Sealed = true
            };

            var dt = new DocumentSetTemplate();
            dt.AllowedContentTypes.Add("0x01005D4F34E4BE7F4B6892AEBE088EDD215E002");
            dt.SharedFields.Add(new Guid("f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3"));
            dt.SharedFields.Add(new Guid("a8df65ec-0d06-4df1-8edf-55d48b3936dc"));
            dt.WelcomePageFields.Add(new Guid("c69d2ffc-0c86-474a-9cc7-dcd7774da531"));
            dt.WelcomePageFields.Add(new Guid("b9132b30-2b9e-47d4-b0fc-1ac34a61506f"));
            dt.WelcomePage = "home.aspx";
            dt.DefaultDocuments.Add(new DefaultDocument()
            {
                ContentTypeId = "0x01005D4F34E4BE7F4B6892AEBE088EDD215E001",
                FileSourcePath = "document.dotx",
                Name = "DefaultDocument"
            });
            nct.DocumentSetTemplate = dt;
            nct.FieldRefs.Add(new FieldRef("TestField")
            {
                Id = new Guid("23203e97-3bfe-40cb-afb4-07aa2b86bf45"),
                Required = true,
                Hidden = true
            });
            nct.FieldRefs.Add(new FieldRef("TestField1"));
            nct.FieldRefs.Add(new FieldRef("TestField2"));
            nct.FieldRefs.Add(new FieldRef("TestField3"));
            result.ContentTypes.Add(nct);

            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-ct.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-ct.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult = 
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.ContentTypes);

            var ct = template.ContentTypes.FirstOrDefault();
            Assert.IsNotNull(ct);

            Assert.AreEqual("0x01005D4F34E4BE7F4B6892AEBE088EDD215E", ct.ID);
            Assert.AreEqual("General Project Document", ct.Name);
            Assert.AreEqual("General Project Document Content Type", ct.Description);
            Assert.AreEqual("Base Foundation Content Types", ct.Group);
            Assert.AreEqual("/Forms/DisplayForm.aspx", ct.DisplayFormUrl);
            Assert.AreEqual("/Forms/NewForm.aspx", ct.NewFormUrl);
            Assert.AreEqual("/Forms/EditForm.aspx", ct.EditFormUrl);
            Assert.IsNotNull(ct.DocumentTemplate);
            Assert.AreEqual("DocumentTemplate.dotx", ct.DocumentTemplate.TargetName);
            Assert.IsTrue(ct.Hidden);
            Assert.IsTrue(ct.Overwrite);
            Assert.IsTrue(ct.ReadOnly);
            Assert.IsTrue(ct.Sealed);

            Assert.IsNotNull(ct.DocumentSetTemplate);
            Assert.IsNotNull(ct.DocumentSetTemplate.AllowedContentTypes);
            Assert.IsNotNull(ct.DocumentSetTemplate.AllowedContentTypes.FirstOrDefault(c => c.ContentTypeID == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E002"));
            Assert.IsNotNull(ct.DocumentSetTemplate.SharedFields);
            Assert.IsNotNull(ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c.ID == "f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3"));
            Assert.IsNotNull(ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c.ID == "a8df65ec-0d06-4df1-8edf-55d48b3936dc"));
            Assert.IsNotNull(ct.DocumentSetTemplate.WelcomePageFields.FirstOrDefault(c => c.ID == "c69d2ffc-0c86-474a-9cc7-dcd7774da531"));
            Assert.IsNotNull(ct.DocumentSetTemplate.WelcomePageFields.FirstOrDefault(c => c.ID == "b9132b30-2b9e-47d4-b0fc-1ac34a61506f"));
            Assert.AreEqual("home.aspx", ct.DocumentSetTemplate.WelcomePage);
            Assert.IsNotNull(ct.DocumentSetTemplate.DefaultDocuments);

            var dd = ct.DocumentSetTemplate.DefaultDocuments.FirstOrDefault(d => d.ContentTypeID == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E001");
            Assert.IsNotNull(dd);
            Assert.AreEqual("document.dotx", dd.FileSourcePath);
            Assert.AreEqual("DefaultDocument", dd.Name);

            Assert.IsNotNull(ct.FieldRefs);
            Assert.AreEqual(4, ct.FieldRefs.Count());

            var field = ct.FieldRefs.FirstOrDefault(f => f.Name == "TestField");
            Assert.IsNotNull(field);
            Assert.AreEqual("23203e97-3bfe-40cb-afb4-07aa2b86bf45", field.ID);
            Assert.IsTrue(field.Required);
            Assert.IsTrue(field.Hidden);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_CustomActions_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.CustomActions);
            Assert.IsNotNull(result.CustomActions.SiteCustomActions);
            Assert.IsNotNull(result.CustomActions.WebCustomActions);

            var ca = result.CustomActions.SiteCustomActions.FirstOrDefault(c => c.Name == "CA_SITE_SETTINGS_SITECLASSIFICATION");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Site Classification Application", ca.Description);
            Assert.AreEqual("Microsoft.SharePoint.SiteSettings", ca.Location);
            Assert.AreEqual("Site Classification", ca.Title);
            Assert.AreEqual(1000, ca.Sequence);
            Assert.IsTrue(ca.Rights.Has(PermissionKind.ManageWeb));
            Assert.AreEqual("https://spmanaged.azurewebsites.net/pages/index.aspx?SPHostUrl={0}", ca.Url);
            Assert.AreEqual(UserCustomActionRegistrationType.None, ca.RegistrationType);
            Assert.IsNotNull(ca.CommandUIExtension);
            Assert.AreEqual("http://sharepoint.com", ca.ImageUrl);
            Assert.AreEqual("101", ca.RegistrationId);
            Assert.AreEqual("alert('boo')", ca.ScriptBlock);
            Assert.AreEqual(2, ca.CommandUIExtension.Nodes().Count());

            ca = result.CustomActions.SiteCustomActions.FirstOrDefault(c => c.Name == "CA_SUBSITE_OVERRIDE");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Override new sub-site link", ca.Description);
            Assert.AreEqual("ScriptLink", ca.Location);
            Assert.AreEqual("SubSiteOveride", ca.Title);
            Assert.AreEqual(100, ca.Sequence);
            Assert.AreEqual("~site/PnP_Provisioning_JS/PnP_EmbeddedJS.js", ca.ScriptSrc);
            Assert.AreEqual(UserCustomActionRegistrationType.ContentType, ca.RegistrationType);
            Assert.IsNull(ca.CommandUIExtension);

            ca = result.CustomActions.WebCustomActions.FirstOrDefault(c => c.Name == "CA_WEB_DOCLIB_MENU_SAMPLE");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Document Library Custom Menu", ca.Description);
            Assert.AreEqual("ActionsMenu", ca.Group);
            Assert.AreEqual("Microsoft.SharePoint.StandardMenu", ca.Location);
            Assert.AreEqual("DocLib Custom Menu", ca.Title);
            Assert.AreEqual(100, ca.Sequence);
            Assert.AreEqual("/_layouts/CustomActionsHello.aspx?ActionsMenu", ca.Url);
            Assert.AreEqual(UserCustomActionRegistrationType.None, ca.RegistrationType);
            Assert.IsNull(ca.CommandUIExtension);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_CustomActions_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = new ProvisioningTemplate();
            result.CustomActions = new CustomActions();

            var can = new CustomAction() {
                Name = "CA_SITE_SETTINGS_SITECLASSIFICATION",
                Description = "Site Classification Application",
                Location = "Microsoft.SharePoint.SiteSettings",
                Title = "Site Classification",
                Url = "https://spmanaged.azurewebsites.net/pages/index.aspx?SPHostUrl={0}",
                Sequence = 1000,
                RegistrationType = UserCustomActionRegistrationType.None,
                Rights = new BasePermissions(),
                ImageUrl = "http://sharepoint.com",
                RegistrationId = "101",
                ScriptBlock = "alert('boo')",
                CommandUIExtension = XElement.Parse(@"<CommandUIExtension><CommandUIDefinitions>
                <CommandUIDefinition Location=""Ribbon.Documents.Copies.Controls._children"">
                  <Button Sequence = ""15"" TemplateAlias = ""o1"" ToolTipDescription = ""Download all files separately"" ToolTipTitle = ""Download All"" Description = ""Download all files separately"" LabelText = ""Download All"" Image32by32 = ""~sitecollection/SiteAssets/DownloadAll32x32.png"" Image16by16 = ""~sitecollection/SiteAssets/DownloadAll16x16.png"" Command = ""OfficeDevPnP.Cmd.DownloadAll"" Id = ""Ribbon.Documents.Copies.OfficeDevPnPDownloadAll"" />
                </CommandUIDefinition>
                <CommandUIDefinition Location = ""Ribbon.Documents.Copies.Controls._children"">
                  <Button Sequence = ""20"" TemplateAlias = ""o1"" ToolTipDescription = ""Download all files as single Zip archive"" ToolTipTitle = ""Download All as Zip"" Description = ""Download all files as single Zip"" LabelText = ""Download All as Zip"" Image32by32 = ""~sitecollection/SiteAssets/DownloadAllAsZip32x32.png"" Image16by16 = ""~sitecollection/SiteAssets/DownloadAllAsZip16x16.png"" Command = ""OfficeDevPnP.Cmd.DownloadAllAsZip"" Id = ""Ribbon.Documents.Copies.OfficeDevPnPDownloadAllAsZip"" />
                </CommandUIDefinition>
              </CommandUIDefinitions>
              <CommandUIHandlers>
                <CommandUIHandler Command = ""OfficeDevPnP.Cmd.DownloadAll"" EnabledScript = ""javascript:OfficeDevPnP.Core.RibbonManager.isListViewButtonEnabled('DownloadAll');"" CommandAction = ""javascript:OfficeDevPnP.Core.RibbonManager.invokeCommand('DownloadAll');"" />
                <CommandUIHandler Command = ""OfficeDevPnP.Cmd.DownloadAllAsZip"" EnabledScript = ""javascript:OfficeDevPnP.Core.RibbonManager.isListViewButtonEnabled('DownloadAllAsZip');"" CommandAction = ""javascript:OfficeDevPnP.Core.RibbonManager.invokeCommand('DownloadAllAsZip');"" />
              </CommandUIHandlers></CommandUIExtension>")
            };
            can.Rights.Set(PermissionKind.ManageWeb);
            result.CustomActions.SiteCustomActions.Add(can);

            can = new CustomAction()
            {
                Name = "CA_SUBSITE_OVERRIDE",
                Description = "Override new sub-site link",
                Location = "ScriptLink",
                Title  = "SubSiteOveride",
                Sequence = 100,
                ScriptSrc = "~site/PnP_Provisioning_JS/PnP_EmbeddedJS.js",
                RegistrationType = UserCustomActionRegistrationType.ContentType
            };
            result.CustomActions.SiteCustomActions.Add(can);

            can = new CustomAction()
            {
                Name = "CA_WEB_DOCLIB_MENU_SAMPLE",
                Description = "Document Library Custom Menu",
                Group = "ActionsMenu",
                Location = "Microsoft.SharePoint.StandardMenu",
                Title = "DocLib Custom Menu",
                Sequence = 100,
                Url = "/_layouts/CustomActionsHello.aspx?ActionsMenu",
                RegistrationType = UserCustomActionRegistrationType.None
            };
            result.CustomActions.WebCustomActions.Add(can);

            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-ca.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-ca.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.CustomActions.SiteCustomActions);
            Assert.IsNotNull(template.CustomActions.WebCustomActions);

            var ca = template.CustomActions.SiteCustomActions.FirstOrDefault(c => c.Name == "CA_SITE_SETTINGS_SITECLASSIFICATION");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Site Classification Application", ca.Description);
            Assert.AreEqual("Microsoft.SharePoint.SiteSettings", ca.Location);
            Assert.AreEqual("Site Classification", ca.Title);
            Assert.AreEqual(1000, ca.Sequence);
            Assert.AreEqual("ManageWeb", ca.Rights);
            Assert.AreEqual("https://spmanaged.azurewebsites.net/pages/index.aspx?SPHostUrl={0}", ca.Url);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.RegistrationType.None, ca.RegistrationType);
            Assert.IsNotNull(ca.CommandUIExtension);
            Assert.AreEqual("http://sharepoint.com", ca.ImageUrl);
            Assert.AreEqual("101", ca.RegistrationId);
            Assert.AreEqual("alert('boo')", ca.ScriptBlock);
            Assert.IsNotNull(ca.CommandUIExtension);
            Assert.IsNotNull(ca.CommandUIExtension.Any);
            Assert.AreEqual(2, ca.CommandUIExtension.Any.Length);

            ca = template.CustomActions.SiteCustomActions.FirstOrDefault(c => c.Name == "CA_SUBSITE_OVERRIDE");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Override new sub-site link", ca.Description);
            Assert.AreEqual("ScriptLink", ca.Location);
            Assert.AreEqual("SubSiteOveride", ca.Title);
            Assert.AreEqual(100, ca.Sequence);
            Assert.AreEqual("~site/PnP_Provisioning_JS/PnP_EmbeddedJS.js", ca.ScriptSrc);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.RegistrationType.ContentType, ca.RegistrationType);
            Assert.IsNull(ca.CommandUIExtension);

            ca = template.CustomActions.WebCustomActions.FirstOrDefault(c => c.Name == "CA_WEB_DOCLIB_MENU_SAMPLE");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Document Library Custom Menu", ca.Description);
            Assert.AreEqual("ActionsMenu", ca.Group);
            Assert.AreEqual("Microsoft.SharePoint.StandardMenu", ca.Location);
            Assert.AreEqual("DocLib Custom Menu", ca.Title);
            Assert.AreEqual(100, ca.Sequence);
            Assert.AreEqual("/_layouts/CustomActionsHello.aspx?ActionsMenu", ca.Url);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.RegistrationType.None, ca.RegistrationType);
            Assert.IsNull(ca.CommandUIExtension);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Files_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.Files);

            var file = result.Files.FirstOrDefault(f => f.Src == "/SitePages/home.aspx");
            Assert.IsNotNull(file);
            Assert.AreEqual("SitePages", file.Folder);
            Assert.AreEqual(Core.Framework.Provisioning.Model.FileLevel.Published, file.Level);
            Assert.IsTrue(file.Overwrite);

            Assert.IsNotNull(file.Properties);
            var property = file.Properties.FirstOrDefault(p => p.Key == "MyProperty1");
            Assert.IsNotNull(property);
            Assert.AreEqual("Value1", property.Value);
            property = file.Properties.FirstOrDefault(p => p.Key == "MyProperty2");
            Assert.IsNotNull(property);
            Assert.AreEqual("Value2", property.Value);

            Assert.IsNotNull(file.Security);
            Assert.IsTrue(file.Security.ClearSubscopes);
            Assert.IsTrue(file.Security.CopyRoleAssignments);
            Assert.IsNotNull(file.Security.RoleAssignments);
            var assingment = file.Security.RoleAssignments.FirstOrDefault(r => r.Principal == "admin@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("Owner", assingment.RoleDefinition);
            assingment = file.Security.RoleAssignments.FirstOrDefault(r => r.Principal == "dev@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("Contributor", assingment.RoleDefinition);

            Assert.IsNotNull(file.WebParts);
            var webpart = file.WebParts.FirstOrDefault(wp => wp.Title == "My Content");
            Assert.IsNotNull(webpart);
            Assert.AreEqual((uint)1, webpart.Order);
            Assert.AreEqual("Main", webpart.Zone);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<Contents xmlns=\"http://schemas.dev.office.com/PnP/2016/05/ProvisioningSchema\"><webParts xmlns=\"\"><webPart>[!<![CDATA[web part definition goes here]]></webPart></webParts></Contents>", webpart.Contents.Trim());

            Assert.IsNotNull(file.WebParts);
            webpart = file.WebParts.FirstOrDefault(wp => wp.Title == "My Editor");
            Assert.IsNotNull(webpart);
            Assert.AreEqual((uint)10, webpart.Order);
            Assert.AreEqual("Left", webpart.Zone);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<Contents xmlns=\"http://schemas.dev.office.com/PnP/2016/05/ProvisioningSchema\"><webParts xmlns=\"\"><webPart>[!<![CDATA[web part definition goes here]]></webPart></webParts></Contents>", webpart.Contents.Trim());

            file = result.Files.FirstOrDefault(f => f.Src == "/Resources/Files/SAMPLE.js");
            Assert.IsNotNull(file);
            Assert.AreEqual("SAMPLE", file.Folder);
            Assert.AreEqual(Core.Framework.Provisioning.Model.FileLevel.Draft, file.Level);
            Assert.IsFalse(file.Overwrite);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Files_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            var newfile = new Core.Framework.Provisioning.Model.File()
            {
                Src = "/SitePages/home.aspx",
                Folder = "SitePages",
                Level = Core.Framework.Provisioning.Model.FileLevel.Published,
                Overwrite = true,
                Security = new ObjectSecurity()
                {
                    ClearSubscopes = true,
                    CopyRoleAssignments = true,
                }
            };
            newfile.Properties.Add("MyProperty1", "Value1");
            newfile.Properties.Add("MyProperty2", "Value2");
            newfile.Security.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment() { Principal = "admin@sharepoint.com", RoleDefinition = "Owner" });
            newfile.Security.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment() { Principal = "dev@sharepoint.com", RoleDefinition = "Contributor" });
            newfile.WebParts.Add(new WebPart() { Title = "My Content", Order = 1, Zone = "Main", Contents = "<webPart>[!<![CDATA[web part definition goes here]]></webPart>" });
            newfile.WebParts.Add(new WebPart() { Title = "My Editor", Order = 10, Zone = "Left", Contents = "<webPart>[!<![CDATA[web part definition goes here]]></webPart>" });
            result.Files.Add(newfile);

            newfile = new Core.Framework.Provisioning.Model.File()
            {
                Src= "/Resources/Files/SAMPLE.js",
                Folder = "SAMPLE",
                Level = Core.Framework.Provisioning.Model.FileLevel.Draft,
                Overwrite = false
            };
            result.Files.Add(newfile);

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-files.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-files.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.Files);
            Assert.IsNotNull(template.Files.File);
            var file = template.Files.File.FirstOrDefault(f => f.Src == "/SitePages/home.aspx");
            Assert.IsNotNull(file);
            Assert.AreEqual("SitePages", file.Folder);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.FileLevel.Published, file.Level);
            Assert.IsTrue(file.Overwrite);

            Assert.IsNotNull(file.Properties);
            var property = file.Properties.FirstOrDefault(p => p.Key == "MyProperty1");
            Assert.IsNotNull(property);
            Assert.AreEqual("Value1", property.Value);
            property = file.Properties.FirstOrDefault(p => p.Key == "MyProperty2");
            Assert.IsNotNull(property);
            Assert.AreEqual("Value2", property.Value);

            Assert.IsNotNull(file.Security);
            Assert.IsNotNull(file.Security.BreakRoleInheritance);
            Assert.IsTrue(file.Security.BreakRoleInheritance.ClearSubscopes);
            Assert.IsTrue(file.Security.BreakRoleInheritance.CopyRoleAssignments);
            Assert.IsNotNull(file.Security.BreakRoleInheritance.RoleAssignment);
            var assingment = file.Security.BreakRoleInheritance.RoleAssignment.FirstOrDefault(r => r.Principal == "admin@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("Owner", assingment.RoleDefinition);
            assingment = file.Security.BreakRoleInheritance.RoleAssignment.FirstOrDefault(r => r.Principal == "dev@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("Contributor", assingment.RoleDefinition);

            Assert.IsNotNull(file.WebParts);
            var webpart = file.WebParts.FirstOrDefault(wp => wp.Title == "My Content");
            Assert.IsNotNull(webpart);
            Assert.AreEqual(1, webpart.Order);
            Assert.AreEqual("Main", webpart.Zone);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents.InnerXml);

            Assert.IsNotNull(file.WebParts);
            webpart = file.WebParts.FirstOrDefault(wp => wp.Title == "My Editor");
            Assert.IsNotNull(webpart);
            Assert.AreEqual(10, webpart.Order);
            Assert.AreEqual("Left", webpart.Zone);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents.InnerXml);

            file = template.Files.File.FirstOrDefault(f => f.Src == "/Resources/Files/SAMPLE.js");
            Assert.IsNotNull(file);
            Assert.AreEqual("SAMPLE", file.Folder);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.FileLevel.Draft, file.Level);
            Assert.IsFalse(file.Overwrite);
            Assert.IsNull(file.Properties);
            Assert.IsNull(file.WebParts);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Directories_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.Directories); new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

             var dir = result.Directories.FirstOrDefault(d => d.Folder == "MyFolder");
            Assert.IsNotNull(dir);
            Assert.AreEqual("SiteAssets", dir.Src);
            Assert.AreEqual(Core.Framework.Provisioning.Model.FileLevel.Published, dir.Level);
            Assert.IsTrue(dir.Overwrite);
            Assert.IsTrue(dir.Recursive);
            Assert.AreEqual(".aspx", dir.ExcludedExtensions);
            Assert.AreEqual(".docx", dir.IncludedExtensions);
            Assert.AreEqual("metafile", dir.MetadataMappingFile);

            Assert.IsNotNull(dir.Security);
            Assert.IsTrue(dir.Security.ClearSubscopes);
            Assert.IsTrue(dir.Security.CopyRoleAssignments);
            Assert.IsNotNull(dir.Security.RoleAssignments);
            var assingment = dir.Security.RoleAssignments.FirstOrDefault(r => r.Principal == "admin@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("owner", assingment.RoleDefinition);
            assingment = dir.Security.RoleAssignments.FirstOrDefault(r => r.Principal == "dev@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("contributor", assingment.RoleDefinition);

            dir = result.Directories.FirstOrDefault(d => d.Folder == "MyFolder2");
            Assert.IsNotNull(dir);
            Assert.AreEqual("Documents", dir.Src);
            Assert.AreEqual(Core.Framework.Provisioning.Model.FileLevel.Checkout, dir.Level);
            Assert.IsFalse(dir.Overwrite);
            Assert.IsFalse(dir.Recursive);
            Assert.AreEqual(".xslx", dir.ExcludedExtensions);
            Assert.AreEqual(".txt", dir.IncludedExtensions);
            Assert.AreEqual("metafile2", dir.MetadataMappingFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Directories_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            var newdir = new Core.Framework.Provisioning.Model.Directory()
            {
                Folder = "MyFolder",
                Level = Core.Framework.Provisioning.Model.FileLevel.Published,
                Overwrite = true,
                Src = "SiteAssets",
                ExcludedExtensions = ".aspx",
                IncludedExtensions = ".docx",
                MetadataMappingFile = "metafile",
                Recursive = true,
                Security = new ObjectSecurity(new List<Core.Framework.Provisioning.Model.RoleAssignment>()
                {
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal = "admin@sharepoint.com",
                        RoleDefinition = "owner"
                    },
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal = "dev@sharepoint.com",
                        RoleDefinition = "contributor"
                    }
                })
                {
                    ClearSubscopes = true,
                    CopyRoleAssignments = true,
                }
            };
            result.Directories.Add(newdir);

            newdir = new Core.Framework.Provisioning.Model.Directory()
            {
                Folder = "MyFolder2",
                Level = Core.Framework.Provisioning.Model.FileLevel.Checkout,
                Overwrite = false,
                Src = "Documents",
                ExcludedExtensions = ".xslx",
                IncludedExtensions = ".txt",
                MetadataMappingFile = "metafile2",
                Recursive = false
            };
            result.Directories.Add(newdir);

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-dirs.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-dirs.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.Files);
            Assert.IsNotNull(template.Files.Directory);
            var dir = template.Files.Directory.FirstOrDefault(d => d.Folder == "MyFolder");
            Assert.IsNotNull(dir);
            Assert.AreEqual("SiteAssets", dir.Src);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.FileLevel.Published, dir.Level);
            Assert.IsTrue(dir.Overwrite);
            Assert.IsTrue(dir.Recursive);
            Assert.AreEqual(".aspx", dir.ExcludedExtensions);
            Assert.AreEqual(".docx", dir.IncludedExtensions);
            Assert.AreEqual("metafile", dir.MetadataMappingFile);

            Assert.IsNotNull(dir.Security);
            Assert.IsNotNull(dir.Security.BreakRoleInheritance);
            Assert.IsTrue(dir.Security.BreakRoleInheritance.ClearSubscopes);
            Assert.IsTrue(dir.Security.BreakRoleInheritance.CopyRoleAssignments);
            Assert.IsNotNull(dir.Security.BreakRoleInheritance.RoleAssignment);
            var assingment = dir.Security.BreakRoleInheritance.RoleAssignment.FirstOrDefault(r => r.Principal == "admin@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("owner", assingment.RoleDefinition);
            assingment = dir.Security.BreakRoleInheritance.RoleAssignment.FirstOrDefault(r => r.Principal == "dev@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("contributor", assingment.RoleDefinition);

            dir = template.Files.Directory.FirstOrDefault(d => d.Folder == "MyFolder2");
            Assert.IsNotNull(dir);
            Assert.AreEqual("Documents", dir.Src);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.FileLevel.Checkout, dir.Level);
            Assert.IsFalse(dir.Overwrite);
            Assert.IsFalse(dir.Recursive);
            Assert.AreEqual(".xslx", dir.ExcludedExtensions);
            Assert.AreEqual(".txt", dir.IncludedExtensions);
            Assert.AreEqual("metafile2", dir.MetadataMappingFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Pages_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.Directories); new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            Assert.IsNotNull(result.Pages);
            var page = result.Pages.FirstOrDefault(d => d.Url == "home.aspx");
            Assert.IsNotNull(page);
            Assert.IsTrue(page.Overwrite);
            Assert.AreEqual(WikiPageLayout.ThreeColumnsHeaderFooter, page.Layout);

            Assert.IsNotNull(page.Security);
            Assert.IsTrue(page.Security.ClearSubscopes);
            Assert.IsTrue(page.Security.CopyRoleAssignments);
            Assert.IsNotNull(page.Security.RoleAssignments);
            var assingment = page.Security.RoleAssignments.FirstOrDefault(r => r.Principal == "admin@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("owner", assingment.RoleDefinition);
            assingment = page.Security.RoleAssignments.FirstOrDefault(r => r.Principal == "dev@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("contributor", assingment.RoleDefinition);

            Assert.IsNotNull(page.WebParts);
            var webpart = page.WebParts.FirstOrDefault(wp => wp.Title == "My Content");
            Assert.IsNotNull(webpart);
            Assert.AreEqual((uint)1, webpart.Row);
            Assert.AreEqual((uint)2, webpart.Column);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webParts><webPart>[!<![CDATA[web part definition goes here]]></webPart></webParts>", webpart.Contents);

            Assert.IsNotNull(page.WebParts);
            webpart = page.WebParts.FirstOrDefault(wp => wp.Title == "My Editor");
            Assert.IsNotNull(webpart);
            Assert.AreEqual((uint)2, webpart.Row);
            Assert.AreEqual((uint)1, webpart.Column);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webParts><webPart>[!<![CDATA[web part definition goes here]]></webPart></webParts>", webpart.Contents);

            Assert.IsNotNull(page.Fields);
            Assert.AreEqual(4, page.Fields.Count() );
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.Key == "TestField"));
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.Key == "TestField2"));
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.Key == "TestField3"));
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.Key == "TestField4"));
            Assert.AreEqual("Value1", page.Fields.FirstOrDefault(f => f.Key == "TestField").Value);
            Assert.AreEqual("Value2", page.Fields.FirstOrDefault(f => f.Key == "TestField2").Value);
            Assert.AreEqual("Value3", page.Fields.FirstOrDefault(f => f.Key == "TestField3").Value);
            Assert.AreEqual("Value4", page.Fields.FirstOrDefault(f => f.Key == "TestField4").Value);

            page = result.Pages.FirstOrDefault(d => d.Url == "help.aspx");
            Assert.IsFalse(page.Overwrite);
            Assert.AreEqual(WikiPageLayout.OneColumnSideBar, page.Layout);
            Assert.IsNull(page.Security);
            Assert.IsTrue(page.WebParts == null || page.WebParts.Count == 0);
            Assert.IsTrue(page.Fields == null || page.Fields.Count == 0);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Pages_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            var newpage = new Core.Framework.Provisioning.Model.Page("home.aspx", true, WikiPageLayout.ThreeColumnsHeaderFooter, new List<WebPart>(), new ObjectSecurity());
            newpage.Security.CopyRoleAssignments = true;
            newpage.Security.ClearSubscopes = true;
            newpage.Security.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment() { Principal = "admin@sharepoint.com", RoleDefinition = "owner" });
            newpage.Security.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment() { Principal = "dev@sharepoint.com", RoleDefinition = "contributor" });
            newpage.WebParts.Add(new WebPart() { Title = "My Content", Row = 2, Column = 1, Contents = "<webPart>[!<![CDATA[web part definition goes here]]></webPart>" });
            newpage.WebParts.Add(new WebPart() { Title = "My Editor", Row = 1, Column = 2, Contents = "<webPart>[!<![CDATA[web part definition goes here]]></webPart>" });
            newpage.Layout = WikiPageLayout.ThreeColumnsHeaderFooter;
            newpage.Fields.Add("TestField", "Value1");
            newpage.Fields.Add("TestField2", "Value2");
            newpage.Fields.Add("TestField3", "Value3");
            newpage.Fields.Add("TestField4", "Value4");
            result.Pages.Add(newpage);

            newpage = new Core.Framework.Provisioning.Model.Page()
            {
                Url = "help.aspx",
                Overwrite = false,
            };
            result.Pages.Add(newpage);

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-pages.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-pages.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.Pages);
            var page = template.Pages.FirstOrDefault(d => d.Url == "home.aspx");
            Assert.IsNotNull(page);
            Assert.IsTrue(page.Overwrite);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.WikiPageLayout.ThreeColumnsHeaderFooter, page.Layout);

            Assert.IsNotNull(page.Security);
            Assert.IsNotNull(page.Security.BreakRoleInheritance);
            Assert.IsTrue(page.Security.BreakRoleInheritance.ClearSubscopes);
            Assert.IsTrue(page.Security.BreakRoleInheritance.CopyRoleAssignments);
            Assert.IsNotNull(page.Security.BreakRoleInheritance.RoleAssignment);
            var assingment = page.Security.BreakRoleInheritance.RoleAssignment.FirstOrDefault(r => r.Principal == "admin@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("owner", assingment.RoleDefinition);
            assingment = page.Security.BreakRoleInheritance.RoleAssignment.FirstOrDefault(r => r.Principal == "dev@sharepoint.com");
            Assert.IsNotNull(assingment);
            Assert.AreEqual("contributor", assingment.RoleDefinition);

            Assert.IsNotNull(page.WebParts);
            var webpart = page.WebParts.FirstOrDefault(wp => wp.Title == "My Content");
            Assert.IsNotNull(webpart);
            Assert.AreEqual(2, webpart.Row);
            Assert.AreEqual(1, webpart.Column);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents.InnerXml);

            Assert.IsNotNull(page.WebParts);
            webpart = page.WebParts.FirstOrDefault(wp => wp.Title == "My Editor");
            Assert.IsNotNull(webpart);
            Assert.AreEqual(1, webpart.Row);
            Assert.AreEqual(2, webpart.Column);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents.InnerXml);

            Assert.IsNotNull(page.Fields);
            Assert.AreEqual(4, page.Fields.Count() );
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.FieldName == "TestField"));
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.FieldName == "TestField2"));
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.FieldName == "TestField3"));
            Assert.IsNotNull(page.Fields.FirstOrDefault(f => f.FieldName == "TestField4"));
            Assert.AreEqual("Value1", page.Fields.FirstOrDefault(f => f.FieldName == "TestField").Value);
            Assert.AreEqual("Value2", page.Fields.FirstOrDefault(f => f.FieldName == "TestField2").Value);
            Assert.AreEqual("Value3", page.Fields.FirstOrDefault(f => f.FieldName == "TestField3").Value);
            Assert.AreEqual("Value4", page.Fields.FirstOrDefault(f => f.FieldName == "TestField4").Value);

            page = template.Pages.FirstOrDefault(d => d.Url == "help.aspx");
            Assert.IsFalse(page.Overwrite);
            Assert.IsNull(page.Security);
            Assert.IsNull(page.WebParts);
            Assert.IsNull(page.Fields);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TermGroups_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.TermGroups);
            Assert.AreEqual(2, result.TermGroups.Count());
            var group = result.TermGroups.FirstOrDefault(d => d.Id == new Guid("21d7d506-1783-4aed-abdf-160f37bd0ca9"));
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Term Group", group.Description);
            Assert.AreEqual("TestTermGroup", group.Name);
            Assert.IsTrue(group.SiteCollectionTermGroup);
            Assert.IsNotNull(group.Contributors);
            Assert.AreEqual(2, group.Contributors.Count());
            Assert.IsNotNull(group.Contributors.FirstOrDefault(u => u.Name == "contributor1@termgroup1"));
            Assert.IsNotNull(group.Contributors.FirstOrDefault(u => u.Name == "contributor2@termgroup1"));
            Assert.IsNotNull(group.Managers);
            Assert.AreEqual(2, group.Managers.Count());
            Assert.IsNotNull(group.Managers.FirstOrDefault(u => u.Name == "manager1@termgroup1"));
            Assert.IsNotNull(group.Managers.FirstOrDefault(u => u.Name == "manager2@termgroup1"));

            Assert.IsNotNull(group.TermSets);
            Assert.AreEqual(2, group.TermSets.Count());
            var ts = group.TermSets.FirstOrDefault(t => t.Id == new Guid("ce70be1b-1772-49e9-a08f-47192d88dd64"));
            Assert.IsNotNull(ts);
            Assert.AreEqual("TestTermset1TestTermGroup", ts.Name);
            Assert.AreEqual("Test Termset 1 Test Term Group", ts.Description);
            Assert.AreEqual("termset1owner@termgroup1", ts.Owner);
            Assert.AreEqual(1049, ts.Language);
            Assert.IsTrue(ts.IsAvailableForTagging);
            Assert.IsTrue(ts.IsOpenForTermCreation);
            Assert.IsNotNull(ts.Properties);
            Assert.AreEqual(2, ts.Properties.Count());
            Assert.IsNotNull(ts.Properties.FirstOrDefault(p => p.Key == "Property1"));
            Assert.IsNotNull(ts.Properties.FirstOrDefault(p => p.Key == "Property2"));
            Assert.AreEqual("Value1", ts.Properties.FirstOrDefault(p => p.Key == "Property1").Value);
            Assert.AreEqual("Value2", ts.Properties.FirstOrDefault(p => p.Key == "Property2").Value);
            Assert.IsNotNull(ts.Terms);
            Assert.AreEqual(2, ts.Terms.Count());

            var tm = ts.Terms.FirstOrDefault(t => t.Id == new Guid("2194b058-c6e0-4805-b875-78cd7d7dfd39"));
            Assert.IsNotNull(tm);
            Assert.AreEqual("Term1Set1Group1", tm.Name);
            Assert.AreEqual("Term1 Set1 Group1", tm.Description);
            Assert.AreEqual(101, tm.CustomSortOrder);
            Assert.AreEqual(1055, tm.Language);
            Assert.AreEqual("term1owner@termgroup1", tm.Owner);
            Assert.AreEqual(new Guid("bd36d6f6-ee5f-4ce5-961c-93867d8f1f3d"), tm.SourceTermId);
            Assert.IsTrue(tm.IsAvailableForTagging);
            Assert.IsTrue(tm.IsDeprecated);
            Assert.IsTrue(tm.IsReused);
            Assert.IsTrue(tm.IsSourceTerm);
            Assert.IsNotNull(tm.LocalProperties);
            Assert.AreEqual(2, tm.LocalProperties.Count());
            Assert.IsNotNull(tm.LocalProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty1"));
            Assert.IsNotNull(tm.LocalProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty2"));
            Assert.AreEqual("Value1", tm.LocalProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty1").Value);
            Assert.AreEqual("Value2", tm.LocalProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty2").Value);
            Assert.IsNotNull(tm.Properties);
            Assert.AreEqual(2, tm.Properties.Count());
            Assert.IsNotNull(tm.Properties.FirstOrDefault(p => p.Key == "Term1Property1"));
            Assert.IsNotNull(tm.Properties.FirstOrDefault(p => p.Key == "Term1Property2"));
            Assert.AreEqual("Value1", tm.Properties.FirstOrDefault(p => p.Key == "Term1Property1").Value);
            Assert.AreEqual("Value2", tm.Properties.FirstOrDefault(p => p.Key == "Term1Property2").Value);
            Assert.IsNotNull(tm.Labels);
            Assert.AreEqual(3, tm.Labels.Count());
            Assert.IsNotNull(tm.Labels.FirstOrDefault(l => l.Language == 1033));
            Assert.AreEqual("Term1Label1033", tm.Labels.FirstOrDefault(l => l.Language == 1033).Value);
            Assert.IsTrue(tm.Labels.FirstOrDefault(l => l.Language == 1033).IsDefaultForLanguage);
            Assert.IsNotNull(tm.Labels.FirstOrDefault(l => l.Language == 1023));
            Assert.AreEqual("Term1Label1023", tm.Labels.FirstOrDefault(l => l.Language == 1023).Value);
            Assert.IsTrue(tm.Labels.FirstOrDefault(l => l.Language == 1023).IsDefaultForLanguage);
            Assert.IsNotNull(tm.Labels.FirstOrDefault(l => l.Language == 1053));
            Assert.AreEqual("Term1Label1023", tm.Labels.FirstOrDefault(l => l.Language == 1053).Value);
            Assert.IsFalse(tm.Labels.FirstOrDefault(l => l.Language == 1053).IsDefaultForLanguage);

            Assert.IsNotNull(tm.Terms);
            Assert.AreEqual(2, tm.Terms.Count());
            var stm = tm.Terms.FirstOrDefault(t => t.Id == new Guid("48fd66cb-f7ca-4160-be46-b78876626c09"));
            Assert.IsNotNull(stm);
            Assert.AreEqual("Subterm1Term1Set1Group1", stm.Name);
            Assert.IsNotNull(stm.Terms);
            Assert.AreEqual(1, stm.Terms.Count());
            Assert.IsNotNull(stm.Terms.FirstOrDefault(t => t.Id == new Guid("7f43fe4a-7030-4d7e-ab62-5fdaac65ac9b")));
            Assert.AreEqual("Subsubterm1Term1Set1Group1", stm.Terms.FirstOrDefault(t => t.Id == new Guid("7f43fe4a-7030-4d7e-ab62-5fdaac65ac9b")).Name);
            stm = tm.Terms.FirstOrDefault(t => t.Id == new Guid("b0d92a3a-cbdf-4c6c-8807-54e23da108ee"));
            Assert.IsNotNull(stm);
            Assert.AreEqual("Subterm2Term1Set1Group1", stm.Name);

            tm = ts.Terms.FirstOrDefault(t => t.Id == new Guid("382d3cb1-89f5-4809-b607-1634698e027e"));
            Assert.IsNotNull(tm);
            Assert.AreEqual("Term2Set1Group1", tm.Name);
            Assert.AreEqual("Term2 Set1 Group1", tm.Description);
            Assert.AreEqual(102, tm.CustomSortOrder);
            Assert.IsNull(tm.Language);
            Assert.AreEqual("term1owner@term2owner", tm.Owner);
            Assert.AreEqual(Guid.Empty, tm.SourceTermId);
            Assert.IsFalse(tm.IsAvailableForTagging);
            Assert.IsFalse(tm.IsDeprecated);
            Assert.IsFalse(tm.IsReused);
            Assert.IsFalse(tm.IsSourceTerm);

            Assert.IsTrue(tm.LocalProperties == null || tm.LocalProperties.Count() == 0);
            Assert.IsTrue(tm.Properties == null || tm.Properties.Count() == 0);
            Assert.IsTrue(tm.Labels == null || tm.Labels.Count() == 0);

            group = result.TermGroups.FirstOrDefault(d => d.Id == new Guid("7d4caedf-4ed3-4e2d-ba93-a166b4f173f6"));
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Term Group 2", group.Description);
            Assert.AreEqual("TestTermGroup2", group.Name);
            Assert.IsFalse(group.SiteCollectionTermGroup);
            Assert.IsTrue(group.TermSets == null || group.TermSets.Count() == 0);
            Assert.IsTrue(group.Contributors == null || group.Contributors.Count() == 0);
            Assert.IsTrue(group.Managers == null || group.Managers.Count() == 0);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TermGroups_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            var termgroup = new TermGroup()
            {
                Id = new Guid("21d7d506-1783-4aed-abdf-160f37bd0ca9"),
                Description = "Test Term Group",
                Name = "TestTermGroup",
                SiteCollectionTermGroup = true
            };
            termgroup.Contributors.Add(new Core.Framework.Provisioning.Model.User() { Name = "contributor1@termgroup1" });
            termgroup.Contributors.Add(new Core.Framework.Provisioning.Model.User() { Name = "contributor2@termgroup1" });
            termgroup.Managers.Add(new Core.Framework.Provisioning.Model.User() { Name = "manager1@termgroup1" });
            termgroup.Managers.Add(new Core.Framework.Provisioning.Model.User() { Name = "manager2@termgroup1" });

#region termset 1 group 1
            var termset = new TermSet()
            {
                Id = new Guid("ce70be1b-1772-49e9-a08f-47192d88dd64"),
                Name = "TestTermset1TestTermGroup",
                Description = "Test Termset 1 Test Term Group",
                IsAvailableForTagging = true,
                IsOpenForTermCreation = true,
                Language = 1049,
                Owner = "termset1owner@termgroup1",
            };

            termset.Properties.Add("Property1", "Value1");
            termset.Properties.Add("Property2", "Value2");

            var term = new Term()
            {
                Id = new Guid("2194b058-c6e0-4805-b875-78cd7d7dfd39"),
                Name = "Term1Set1Group1",
                Description = "Term1 Set1 Group1",
                CustomSortOrder = 101,
                IsAvailableForTagging = true,
                IsDeprecated = true,
                IsReused = true,
                IsSourceTerm = true,
                Language = 1055,
                Owner = "term1owner@termgroup1",
                SourceTermId = new Guid("bd36d6f6-ee5f-4ce5-961c-93867d8f1f3d"),
                
            };
            term.LocalProperties.Add("Term1LocalProperty1", "Value1");
            term.LocalProperties.Add("Term1LocalProperty2", "Value2");
            term.Properties.Add("Term1Property1", "Value1");
            term.Properties.Add("Term1Property2", "Value2");

            term.Labels.Add(new TermLabel() { IsDefaultForLanguage = true, Language = 1033, Value = "Term1Label1033" });
            term.Labels.Add(new TermLabel() { IsDefaultForLanguage = true, Language = 1023, Value = "Term1Label1023" });
            term.Labels.Add(new TermLabel() { IsDefaultForLanguage = false, Language = 1053, Value = "Term1Label1053" });

            var subterm = new Term()
            {
                Id = new Guid("48fd66cb-f7ca-4160-be46-b78876626c09"),
                Name = "Subterm1Term1Set1Group1"
            };

            subterm.Terms.Add(new Term()
            {
                Id = new Guid("7f43fe4a-7030-4d7e-ab62-5fdaac65ac9b"),
                Name = "Subsubterm1Term1Set1Group1"
            });

            term.Terms.Add(subterm);
            term.Terms.Add(new Term()
            {
                Id = new Guid("b0d92a3a-cbdf-4c6c-8807-54e23da108ee"),
                Name = "Subterm2Term1Set1Group1"
            });
            termset.Terms.Add(term);
            termset.Terms.Add(new Term()
            {
                Id = new Guid("382d3cb1-89f5-4809-b607-1634698e027e"),
                Name = "Term2Set1Group1",
                Description = "Term2 Set1 Group1",
                CustomSortOrder = 102,
                IsAvailableForTagging = false,
                IsDeprecated = false,
                IsReused = false,
                IsSourceTerm = false,
                Owner = "term2owner@termgroup1"
            });
            termgroup.TermSets.Add(termset);
#endregion
#region termset 2 group 1
            termset = new TermSet()
            {
                Id = new Guid("d0610999-539c-4949-ba60-0375deea3023"),
                Name = "TestTermset2TestTermGroup",
                Description = "Test Termset 2 Test Term Group",
                IsAvailableForTagging = false,
                IsOpenForTermCreation = false,
            };
            termgroup.TermSets.Add(termset);
#endregion
            result.TermGroups.Add(termgroup);
            termgroup = new TermGroup()
            {
                Id = new Guid("7d4caedf-4ed3-4e2d-ba93-a166b4f173f6"),
                Description = "Test Term Group 2",
                Name = "TestTermGroup2",
                SiteCollectionTermGroup = false
            };
            result.TermGroups.Add(termgroup);

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-tax.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-tax.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.TermGroups);
            Assert.AreEqual(2, template.TermGroups.Count());
            var group = template.TermGroups.FirstOrDefault(d => d.ID == "21d7d506-1783-4aed-abdf-160f37bd0ca9");
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Term Group", group.Description);
            Assert.AreEqual("TestTermGroup", group.Name);
            Assert.IsTrue(group.SiteCollectionTermGroup);
            Assert.IsNotNull(group.Contributors);
            Assert.AreEqual(2, group.Contributors.Count());
            Assert.IsNotNull(group.Contributors.FirstOrDefault(u=>u.Name == "contributor1@termgroup1"));
            Assert.IsNotNull(group.Contributors.FirstOrDefault(u => u.Name == "contributor2@termgroup1"));
            Assert.IsNotNull(group.Managers);
            Assert.AreEqual(2, group.Managers.Count());
            Assert.IsNotNull(group.Managers.FirstOrDefault(u => u.Name == "manager1@termgroup1"));
            Assert.IsNotNull(group.Managers.FirstOrDefault(u => u.Name == "manager2@termgroup1"));

            Assert.IsNotNull(group.TermSets);
            Assert.AreEqual(2, group.TermSets.Count());
            var ts = group.TermSets.FirstOrDefault(t => t.ID == "ce70be1b-1772-49e9-a08f-47192d88dd64");
            Assert.IsNotNull(ts);
            Assert.AreEqual("TestTermset1TestTermGroup", ts.Name);
            Assert.AreEqual("Test Termset 1 Test Term Group", ts.Description);
            Assert.AreEqual("termset1owner@termgroup1", ts.Owner);
            Assert.AreEqual(1049, ts.Language);
            Assert.IsTrue(ts.IsAvailableForTagging);
            Assert.IsTrue(ts.IsOpenForTermCreation);
            Assert.IsNotNull(ts.CustomProperties);
            Assert.AreEqual(2, ts.CustomProperties.Count());
            Assert.IsNotNull(ts.CustomProperties.FirstOrDefault(p => p.Key == "Property1"));
            Assert.IsNotNull(ts.CustomProperties.FirstOrDefault(p => p.Key == "Property2"));
            Assert.AreEqual("Value1", ts.CustomProperties.FirstOrDefault(p => p.Key == "Property1").Value);
            Assert.AreEqual("Value2", ts.CustomProperties.FirstOrDefault(p => p.Key == "Property2").Value);
            Assert.IsNotNull(ts.Terms);
            Assert.AreEqual(2, ts.Terms.Count());

            var tm = ts.Terms.FirstOrDefault(t => t.ID == "2194b058-c6e0-4805-b875-78cd7d7dfd39");
            Assert.IsNotNull(tm);
            Assert.AreEqual("Term1Set1Group1", tm.Name);
            Assert.AreEqual("Term1 Set1 Group1", tm.Description);
            Assert.AreEqual(101, tm.CustomSortOrder);
            Assert.AreEqual(1055, tm.Language);
            Assert.AreEqual("term1owner@termgroup1", tm.Owner);
            Assert.AreEqual("bd36d6f6-ee5f-4ce5-961c-93867d8f1f3d", tm.SourceTermId);
            Assert.IsTrue(tm.IsAvailableForTagging);
            Assert.IsTrue(tm.IsDeprecated);
            Assert.IsTrue(tm.IsReused);
            Assert.IsTrue(tm.IsSourceTerm);
            Assert.IsNotNull(tm.LocalCustomProperties);
            Assert.AreEqual(2, tm.LocalCustomProperties.Count());
            Assert.IsNotNull(tm.LocalCustomProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty1"));
            Assert.IsNotNull(tm.LocalCustomProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty2"));
            Assert.AreEqual("Value1", tm.LocalCustomProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty1").Value);
            Assert.AreEqual("Value2", tm.LocalCustomProperties.FirstOrDefault(p => p.Key == "Term1LocalProperty2").Value);
            Assert.IsNotNull(tm.CustomProperties);
            Assert.AreEqual(2, tm.CustomProperties.Count());
            Assert.IsNotNull(tm.CustomProperties.FirstOrDefault(p => p.Key == "Term1Property1"));
            Assert.IsNotNull(tm.CustomProperties.FirstOrDefault(p => p.Key == "Term1Property2"));
            Assert.AreEqual("Value1", tm.CustomProperties.FirstOrDefault(p => p.Key == "Term1Property1").Value);
            Assert.AreEqual("Value2", tm.CustomProperties.FirstOrDefault(p => p.Key == "Term1Property2").Value);
            Assert.IsNotNull(tm.Labels);
            Assert.AreEqual(3, tm.Labels.Count());
            Assert.IsNotNull(tm.Labels.FirstOrDefault(l => l.Language == 1033));
            Assert.AreEqual("Term1Label1033", tm.Labels.FirstOrDefault(l => l.Language == 1033).Value);
            Assert.IsTrue(tm.Labels.FirstOrDefault(l => l.Language == 1033).IsDefaultForLanguage);
            Assert.IsNotNull(tm.Labels.FirstOrDefault(l => l.Language == 1023));
            Assert.AreEqual("Term1Label1023", tm.Labels.FirstOrDefault(l => l.Language == 1023).Value);
            Assert.IsTrue(tm.Labels.FirstOrDefault(l => l.Language == 1023).IsDefaultForLanguage);
            Assert.IsNotNull(tm.Labels.FirstOrDefault(l => l.Language == 1053));
            Assert.AreEqual("Term1Label1053", tm.Labels.FirstOrDefault(l => l.Language == 1053).Value);
            Assert.IsFalse(tm.Labels.FirstOrDefault(l => l.Language == 1053).IsDefaultForLanguage);

            Assert.IsNotNull(tm.Terms);
            Assert.IsNotNull(tm.Terms.Items);
            Assert.AreEqual(2, tm.Terms.Items.Count());
            var stm = tm.Terms.Items.FirstOrDefault(t => t.ID == "48fd66cb-f7ca-4160-be46-b78876626c09");
            Assert.IsNotNull(stm);
            Assert.AreEqual("Subterm1Term1Set1Group1", stm.Name);
            Assert.IsNotNull(stm.Terms);
            Assert.IsNotNull(stm.Terms.Items);
            Assert.AreEqual(1, stm.Terms.Items.Count());
            Assert.IsNotNull(stm.Terms.Items.FirstOrDefault(t => t.ID == "7f43fe4a-7030-4d7e-ab62-5fdaac65ac9b"));
            Assert.AreEqual("Subsubterm1Term1Set1Group1", stm.Terms.Items.FirstOrDefault(t => t.ID == "7f43fe4a-7030-4d7e-ab62-5fdaac65ac9b").Name);
            stm = tm.Terms.Items.FirstOrDefault(t => t.ID == "b0d92a3a-cbdf-4c6c-8807-54e23da108ee");
            Assert.IsNotNull(stm);
            Assert.AreEqual("Subterm2Term1Set1Group1", stm.Name);

            tm = ts.Terms.FirstOrDefault(t => t.ID == "382d3cb1-89f5-4809-b607-1634698e027e");
            Assert.IsNotNull(tm);
            Assert.AreEqual("Term2Set1Group1", tm.Name);
            Assert.AreEqual("Term2 Set1 Group1", tm.Description);
            Assert.AreEqual(102, tm.CustomSortOrder);
            Assert.IsFalse(tm.LanguageSpecified);
            Assert.AreEqual("term2owner@termgroup1", tm.Owner);
            Assert.IsNull(tm.SourceTermId);
            Assert.IsFalse(tm.IsAvailableForTagging);
            Assert.IsFalse(tm.IsDeprecated);
            Assert.IsFalse(tm.IsReused);
            Assert.IsFalse(tm.IsSourceTerm);

            Assert.IsNull(tm.LocalCustomProperties);
            Assert.IsNull(tm.CustomProperties);
            Assert.IsNull(tm.Labels);

            group = template.TermGroups.FirstOrDefault(d => d.ID == "7d4caedf-4ed3-4e2d-ba93-a166b4f173f6");
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Term Group 2", group.Description);
            Assert.AreEqual("TestTermGroup2", group.Name);
            Assert.IsFalse(group.SiteCollectionTermGroup);
            Assert.IsNotNull(group.TermSets);
            Assert.AreEqual(0, group.TermSets.Length);
            Assert.IsNull(group.Contributors);
            Assert.IsNull(group.Managers);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ComposedLook_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var result = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.ComposedLook);
            Assert.AreEqual("Resources/Themes/Contoso/contosobg.jpg", result.ComposedLook.BackgroundFile);
            Assert.AreEqual("Resources/Themes/Contoso/contoso.spcolor", result.ComposedLook.ColorFile);
            Assert.AreEqual("Resources/Themes/Contoso/contoso.spfont", result.ComposedLook.FontFile);
            Assert.AreEqual("Contoso", result.ComposedLook.Name);
            Assert.AreEqual(2, result.ComposedLook.Version);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ComposedLook_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            var composedLook = new ComposedLook()
            {
                BackgroundFile = "Resources/Themes/Contoso/contosobg.jpg",
                ColorFile = "Resources/Themes/Contoso/contoso.spcolor",
                FontFile = "Resources/Themes/Contoso/contoso.spfont",
                Name = "Contoso",
                Version = 2
            };
            result.ComposedLook = composedLook;

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-look.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-look.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.ComposedLook);
            Assert.AreEqual("Resources/Themes/Contoso/contosobg.jpg", template.ComposedLook.BackgroundFile);
            Assert.AreEqual("Resources/Themes/Contoso/contoso.spcolor", template.ComposedLook.ColorFile);
            Assert.AreEqual("Resources/Themes/Contoso/contoso.spfont", template.ComposedLook.FontFile);
            Assert.AreEqual("Contoso", template.ComposedLook.Name);
            Assert.AreEqual(2, template.ComposedLook.Version);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Workflows_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(template.Workflows);
            Assert.IsNotNull(template.Workflows.WorkflowDefinitions);
            var wd = template.Workflows.WorkflowDefinitions.FirstOrDefault(d => d.Id == new Guid("8fd9de8b-d786-43bf-9b33-d7266eb241b0"));
            Assert.IsNotNull(wd);
            Assert.AreEqual("/workflow1/associate.aspx", wd.AssociationUrl);
            Assert.AreEqual("Test Workflow Definition", wd.Description);
            Assert.AreEqual("My Workflow 1", wd.DisplayName);
            Assert.AreEqual("1.0", wd.DraftVersion);
            Assert.AreEqual("<Field></Field>", wd.FormField);
            Assert.AreEqual("/workflow1/initiate.aspx", wd.InitiationUrl);
            Assert.IsTrue(wd.Published);
            Assert.IsTrue(wd.RequiresAssociationForm);
            Assert.IsTrue(wd.RequiresInitiationForm);
            Assert.AreEqual("List", wd.RestrictToScope);
            Assert.AreEqual("Universal", wd.RestrictToType);
            Assert.AreEqual("workflow1.xaml", wd.XamlPath);
            Assert.IsNotNull(wd.Properties);
            Assert.AreEqual(2, wd.Properties.Count());
            Assert.IsNotNull(wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty1"));
            Assert.AreEqual("Value1", wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty1").Value);
            Assert.IsNotNull(wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty2"));
            Assert.AreEqual("Value2", wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty2").Value);

            wd = template.Workflows.WorkflowDefinitions.FirstOrDefault(d => d.Id == new Guid("13d4bae2-2292-4297-84c5-d56881c529a9"));
            Assert.IsNotNull(wd);
            Assert.IsNull(wd.AssociationUrl);
            Assert.IsNull(wd.Description);
            Assert.AreEqual("My Workflow 2", wd.DisplayName);
            Assert.IsNull(wd.DraftVersion);
            Assert.IsNull(wd.FormField);
            Assert.IsNull(wd.InitiationUrl);
            Assert.IsFalse(wd.Published);
            Assert.IsFalse(wd.RequiresAssociationForm);
            Assert.IsFalse(wd.RequiresInitiationForm);
            Assert.IsNull(wd.RestrictToScope);
            Assert.AreEqual("Universal", wd.RestrictToType);
            Assert.IsTrue(wd.Properties == null || wd.Properties.Count == 0);
            Assert.AreEqual("workflow2.xaml", wd.XamlPath);

            var ws = template.Workflows.WorkflowSubscriptions.FirstOrDefault(d => d.DefinitionId == new Guid("c421e3cb-e7b0-489c-b7cc-e0d35d1179e0"));
            Assert.IsNotNull(ws);
            Assert.IsTrue(ws.Enabled);
            Assert.AreEqual("aa0e4ccf-6f34-4b83-94a4-7b1f28dcf7b7", ws.EventSourceId);
            Assert.IsNotNull(ws.EventTypes);
            Assert.AreEqual(3, ws.EventTypes.Count);
            Assert.IsTrue(ws.EventTypes.Contains("ItemAdded"));
            Assert.IsTrue(ws.EventTypes.Contains("ItemUpdated"));
            Assert.IsTrue(ws.EventTypes.Contains("WorkflowStart"));
            Assert.IsTrue(ws.ManualStartBypassesActivationLimit);
            Assert.AreEqual("94413de1-850d-4fbf-a8bb-371feefa2ecf", ws.ListId);
            Assert.AreEqual("MyWorkflowSubscription1", ws.Name);
            Assert.AreEqual("0x01", ws.ParentContentTypeId);
            Assert.AreEqual("MyWorkflow1Status", ws.StatusFieldName);

            ws = template.Workflows.WorkflowSubscriptions.FirstOrDefault(d => d.DefinitionId == new Guid("34ae3873-3f8e-41b0-aaab-802fc6199897"));
            Assert.IsNotNull(ws);
            Assert.IsFalse(ws.Enabled);
            Assert.IsNull(ws.EventSourceId);
            Assert.IsTrue(ws.EventTypes == null || ws.EventTypes.Count == 0);
            Assert.IsFalse(ws.ManualStartBypassesActivationLimit);
            Assert.IsNull(ws.ListId);
            Assert.AreEqual("MyWorkflowSubscription2", ws.Name);
            Assert.IsNull(ws.ParentContentTypeId);
            Assert.AreEqual("MyWorkflow2Status", ws.StatusFieldName);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Workflows_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            var workflows = new Workflows();
            workflows.WorkflowDefinitions.Add(new WorkflowDefinition(new Dictionary<string, string>() { { "MyWorkflowProperty1", "Value1" }, { "MyWorkflowProperty2", "Value2" } })
            {
                Id = new Guid("8fd9de8b-d786-43bf-9b33-d7266eb241b0"),
                AssociationUrl = "/workflow1/associate.aspx",
                Description = "Test Workflow Definition",
                DisplayName = "My Workflow 1",
                DraftVersion = "1.0",
                FormField = "<Field></Field>",
                InitiationUrl = "/workflow1/initiate.aspx",
                Published = true,
                RequiresAssociationForm = true,
                RequiresInitiationForm = true,
                RestrictToScope = "List",
                RestrictToType = "List",
                XamlPath = "workflow1.xaml"
            });

            workflows.WorkflowDefinitions.Add(new WorkflowDefinition()
            {
                Id = new Guid("13d4bae2-2292-4297-84c5-d56881c529a9"),
                DisplayName = "My Workflow 2",
                XamlPath = "workflow2.xaml"
            });

            workflows.WorkflowSubscriptions.Add(new WorkflowSubscription(new Dictionary<string, string>() { { "MySubscriptionProperty1", "Value1" }, { "MySubscriptionProperty2", "Value2" } })
            {
                DefinitionId = new Guid("c421e3cb-e7b0-489c-b7cc-e0d35d1179e0"),
                Enabled = true,
                EventSourceId = "aa0e4ccf-6f34-4b83-94a4-7b1f28dcf7b7",
                EventTypes = new List<string>() { "ItemAdded", "ItemUpdated", "WorkflowStart" },
                ListId = "94413de1-850d-4fbf-a8bb-371feefa2ecf",
                ManualStartBypassesActivationLimit = true,
                Name = "MyWorkflowSubscription1",
                ParentContentTypeId = "0x01",
                StatusFieldName = "MyWorkflow1Status"
            });
            workflows.WorkflowSubscriptions.Add(new WorkflowSubscription()
            {
                DefinitionId = new Guid("34ae3873-3f8e-41b0-aaab-802fc6199897"),
                Enabled = false,
                Name = "MyWorkflowSubscription2",
                StatusFieldName = "MyWorkflow2Status"
            });

            result.Workflows = workflows;

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-wf.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-wf.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.Workflows);
            Assert.IsNotNull(template.Workflows.WorkflowDefinitions);
            var wd = template.Workflows.WorkflowDefinitions.FirstOrDefault(d => d.Id == "8fd9de8b-d786-43bf-9b33-d7266eb241b0");
            Assert.IsNotNull(wd);
            Assert.AreEqual("/workflow1/associate.aspx", wd.AssociationUrl);
            Assert.AreEqual("Test Workflow Definition", wd.Description);
            Assert.AreEqual("My Workflow 1", wd.DisplayName);
            Assert.AreEqual("1.0", wd.DraftVersion);
            Assert.IsNotNull(wd.FormField);
            Assert.AreEqual("<Field></Field>", wd.FormField.OuterXml);
            Assert.AreEqual("/workflow1/initiate.aspx", wd.InitiationUrl);
            Assert.IsTrue(wd.Published);
            Assert.IsTrue(wd.PublishedSpecified);
            Assert.IsTrue(wd.RequiresAssociationForm);
            Assert.IsTrue(wd.RequiresAssociationFormSpecified);
            Assert.IsTrue(wd.RequiresInitiationForm);
            Assert.IsTrue(wd.RequiresInitiationFormSpecified);
            Assert.AreEqual("List", wd.RestrictToScope);
            Assert.AreEqual(OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605.WorkflowsWorkflowDefinitionRestrictToType.List, wd.RestrictToType);
            Assert.AreEqual("workflow1.xaml", wd.XamlPath);
            Assert.IsNotNull(wd.Properties);
            Assert.AreEqual(2, wd.Properties.Count());
            Assert.IsNotNull(wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty1"));
            Assert.AreEqual("Value1", wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty1").Value);
            Assert.IsNotNull(wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty2"));
            Assert.AreEqual("Value2", wd.Properties.FirstOrDefault(p => p.Key == "MyWorkflowProperty2").Value);

            wd = template.Workflows.WorkflowDefinitions.FirstOrDefault(d => d.Id == "13d4bae2-2292-4297-84c5-d56881c529a9");
            Assert.IsNotNull(wd);
            Assert.IsNull(wd.AssociationUrl);
            Assert.IsNull(wd.Description);
            Assert.AreEqual("My Workflow 2", wd.DisplayName);
            Assert.IsNull(wd.DraftVersion);
            Assert.IsNull(wd.FormField);
            Assert.IsNull(wd.InitiationUrl);
            Assert.IsFalse(wd.Published);
            Assert.IsFalse(wd.PublishedSpecified);
            Assert.IsFalse(wd.RequiresAssociationForm);
            Assert.IsFalse(wd.RequiresAssociationFormSpecified);
            Assert.IsFalse(wd.RequiresInitiationForm);
            Assert.IsFalse(wd.RequiresInitiationFormSpecified);
            Assert.IsNull(wd.RestrictToScope);
            Assert.IsFalse(wd.RestrictToTypeSpecified);
            Assert.IsNull(wd.Properties);
            Assert.AreEqual("workflow2.xaml", wd.XamlPath);

            var ws = template.Workflows.WorkflowSubscriptions.FirstOrDefault(d => d.DefinitionId == "c421e3cb-e7b0-489c-b7cc-e0d35d1179e0");
            Assert.IsNotNull(ws);
            Assert.IsTrue(ws.Enabled);
            Assert.AreEqual("aa0e4ccf-6f34-4b83-94a4-7b1f28dcf7b7", ws.EventSourceId);
            Assert.IsTrue(ws.ItemAddedEvent);
            Assert.IsTrue(ws.ItemUpdatedEvent);
            Assert.IsTrue(ws.WorkflowStartEvent);
            Assert.IsTrue(ws.ManualStartBypassesActivationLimit);
            Assert.IsTrue(ws.ManualStartBypassesActivationLimitSpecified);
            Assert.AreEqual("94413de1-850d-4fbf-a8bb-371feefa2ecf", ws.ListId);
            Assert.AreEqual("MyWorkflowSubscription1", ws.Name);
            Assert.AreEqual("0x01", ws.ParentContentTypeId);
            Assert.AreEqual("MyWorkflow1Status", ws.StatusFieldName);

            ws = template.Workflows.WorkflowSubscriptions.FirstOrDefault(d => d.DefinitionId == "34ae3873-3f8e-41b0-aaab-802fc6199897");
            Assert.IsNotNull(ws);
            Assert.IsFalse(ws.Enabled);
            Assert.IsNull(ws.EventSourceId);
            Assert.IsFalse(ws.ItemAddedEvent);
            Assert.IsFalse(ws.ItemUpdatedEvent);
            Assert.IsFalse(ws.WorkflowStartEvent);
            Assert.IsFalse(ws.ManualStartBypassesActivationLimit);
            Assert.IsFalse(ws.ManualStartBypassesActivationLimitSpecified);
            Assert.IsNull(ws.ListId);
            Assert.AreEqual("MyWorkflowSubscription2", ws.Name);
            Assert.IsNull(ws.ParentContentTypeId);
            Assert.AreEqual("MyWorkflow2Status", ws.StatusFieldName);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SearchSettings_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.AreEqual("<SiteSearchSettings></SiteSearchSettings>", template.SiteSearchSettings);
            Assert.AreEqual("<WebSearchSettings></WebSearchSettings>", template.WebSearchSettings);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SearchSettings_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.SiteSearchSettings = "<SiteSearchSettings></SiteSearchSettings>";
            result.WebSearchSettings = "<WebSearchSettings></WebSearchSettings>";
            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-srch.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-srch.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.SearchSettings);
            Assert.IsNotNull(template.SearchSettings.SiteSearchSettings);
            Assert.IsNotNull(template.SearchSettings.WebSearchSettings);
            Assert.AreEqual("<SiteSearchSettings></SiteSearchSettings>", template.SearchSettings.SiteSearchSettings.OuterXml);
            Assert.AreEqual("<WebSearchSettings></WebSearchSettings>", template.SearchSettings.WebSearchSettings.OuterXml);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Publishing_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(template.Publishing);
            Assert.AreEqual(AutoCheckRequirementsOptions.SkipIfNotCompliant, template.Publishing.AutoCheckRequirements);
            Assert.IsNotNull(template.Publishing.DesignPackage);
            Assert.AreEqual("mypackage", template.Publishing.DesignPackage.DesignPackagePath);
            Assert.AreEqual(2, template.Publishing.DesignPackage.MajorVersion);
            Assert.AreEqual(3, template.Publishing.DesignPackage.MinorVersion);
            Assert.AreEqual(new Guid("306ab10d-981d-471d-a8f9-16e1260ad4eb"), template.Publishing.DesignPackage.PackageGuid);
            Assert.AreEqual("MyTestPackage", template.Publishing.DesignPackage.PackageName);

            Assert.IsNotNull(template.Publishing.AvailableWebTemplates);
            Assert.AreEqual(2, template.Publishing.AvailableWebTemplates.Count());
            Assert.IsNotNull(template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 1033));
            Assert.AreEqual("Template1033", template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 1033).TemplateName);
            Assert.IsNotNull(template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 1049));
            Assert.AreEqual("Template1049", template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 1049).TemplateName);

            Assert.IsNotNull(template.Publishing.PageLayouts);
            Assert.IsNotNull(template.Publishing.PageLayouts);
            Assert.AreEqual(2, template.Publishing.PageLayouts.Count());
            Assert.IsNotNull(template.Publishing.PageLayouts.FirstOrDefault(p => p.Path == "mypagelayout1.aspx"));
            Assert.IsTrue(template.Publishing.PageLayouts.FirstOrDefault(p => p.Path == "mypagelayout1.aspx").IsDefault);
            Assert.IsNotNull(template.Publishing.PageLayouts.FirstOrDefault(p => p.Path == "mypagelayout2.aspx"));
            Assert.IsFalse(template.Publishing.PageLayouts.FirstOrDefault(p => p.Path == "mypagelayout2.aspx").IsDefault);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Publishing_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.Publishing = new Publishing()
            {
                AutoCheckRequirements = AutoCheckRequirementsOptions.SkipIfNotCompliant,
                DesignPackage = new DesignPackage() {
                    DesignPackagePath ="mypackage",
                    MajorVersion = 2,
                    MinorVersion = 3,
                    PackageGuid = new Guid("306ab10d-981d-471d-a8f9-16e1260ad4eb"),
                    PackageName = "MyTestPackage"
                }
            };
            result.Publishing.AvailableWebTemplates.Add(new AvailableWebTemplate()
            {
                TemplateName = "Template1033"
            });
            result.Publishing.AvailableWebTemplates.Add(new AvailableWebTemplate()
            {
                LanguageCode = 1049,
                TemplateName = "Template1049"
            });
            result.Publishing.PageLayouts.Add(new PageLayout()
            {
                IsDefault = true,
                Path = "mypagelayout1.aspx"
            });
            result.Publishing.PageLayouts.Add(new PageLayout()
            {
                IsDefault = false,
                Path = "mypagelayout2.aspx"
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-pub.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-pub.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.Publishing);
            Assert.AreEqual(OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.V201605.PublishingAutoCheckRequirements.SkipIfNotCompliant, template.Publishing.AutoCheckRequirements);
            Assert.IsNotNull(template.Publishing.DesignPackage);
            Assert.AreEqual("mypackage", template.Publishing.DesignPackage.DesignPackagePath);
            Assert.AreEqual(2, template.Publishing.DesignPackage.MajorVersion);
            Assert.IsTrue(template.Publishing.DesignPackage.MajorVersionSpecified);
            Assert.AreEqual(3, template.Publishing.DesignPackage.MinorVersion);
            Assert.IsTrue(template.Publishing.DesignPackage.MinorVersionSpecified);
            Assert.AreEqual("306ab10d-981d-471d-a8f9-16e1260ad4eb", template.Publishing.DesignPackage.PackageGuid);
            Assert.AreEqual("MyTestPackage", template.Publishing.DesignPackage.PackageName);

            Assert.IsNotNull(template.Publishing.AvailableWebTemplates);
            Assert.AreEqual(2, template.Publishing.AvailableWebTemplates.Count());
            Assert.IsNotNull(template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 0));
            Assert.AreEqual("Template1033", template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 0).TemplateName);
            Assert.IsNotNull(template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 1049));
            Assert.AreEqual("Template1049", template.Publishing.AvailableWebTemplates.FirstOrDefault(t => t.LanguageCode == 1049).TemplateName);

            Assert.IsNotNull(template.Publishing.PageLayouts);
            Assert.AreEqual("mypagelayout1.aspx", template.Publishing.PageLayouts.Default);
            Assert.IsNotNull(template.Publishing.PageLayouts.PageLayout);
            Assert.AreEqual(2, template.Publishing.PageLayouts.PageLayout.Count());
            Assert.IsNotNull(template.Publishing.PageLayouts.PageLayout.FirstOrDefault(p => p.Path == "mypagelayout1.aspx"));
            Assert.IsNotNull(template.Publishing.PageLayouts.PageLayout.FirstOrDefault(p => p.Path == "mypagelayout2.aspx"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_AddIns_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(template.AddIns);
            Assert.AreEqual(2, template.AddIns.Count());
            Assert.IsNotNull(template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin1.app"));
            Assert.AreEqual("DeveloperSite", template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin1.app").Source);
            Assert.IsNotNull(template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin2.app"));
            Assert.AreEqual("Marketplace", template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin2.app").Source);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_AddIns_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.AddIns.Add(new AddIn()
            {
                PackagePath = "myaddin1.app",
                Source = "DeveloperSite"
            });

            result.AddIns.Add(new AddIn()
            {
                PackagePath = "myaddin2.app",
                Source = "Marketplace"
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-addin.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-addin.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.AddIns);
            Assert.AreEqual(2, template.AddIns.Count());
            Assert.IsNotNull(template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin1.app"));
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.AddInsAddinSource.DeveloperSite, template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin1.app").Source);
            Assert.IsNotNull(template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin2.app"));
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.AddInsAddinSource.Marketplace, template.AddIns.FirstOrDefault(a => a.PackagePath == "myaddin2.app").Source);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ExtensibilityHandlers_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.ExtensibilityHandlers);
            Assert.AreEqual(2, template.ExtensibilityHandlers.Count());
            var handler = template.ExtensibilityHandlers.FirstOrDefault(p => p.Type == "System.Guid");
            Assert.IsNotNull(handler);
            Assert.AreEqual("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089", handler.Assembly);
            Assert.IsTrue(handler.Enabled);
            Assert.AreEqual("<TestConfiguration xmlns=\"MyHandler\">Value</TestConfiguration>", handler.Configuration.Trim());

            handler = template.ExtensibilityHandlers.FirstOrDefault(p => p.Type == "System.String");
            Assert.IsNotNull(handler);
            Assert.AreEqual("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089", handler.Assembly);
            Assert.IsFalse(handler.Enabled);
            Assert.IsNull(handler.Configuration);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ExtensibilityHandlers_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.ExtensibilityHandlers.Add(new ExtensibilityHandler()
            {
                Type = "MyType",
                Assembly = "MyAssembly",
                Enabled = true,
                Configuration = "<TestConfiguration>Value</TestConfiguration>"

            });

            result.ExtensibilityHandlers.Add(new ExtensibilityHandler()
            {
                Type = "MyType2",
                Assembly = "MyAssembly2",
                Enabled = false
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-addin.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-addin.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.Providers);
            Assert.AreEqual(2, template.Providers.Count());
            var handler = template.Providers.FirstOrDefault(p => p.HandlerType == "MyType, MyAssembly");
            Assert.IsNotNull(handler);
            Assert.IsTrue(handler.Enabled);
            Assert.IsNotNull(handler.Configuration);
            Assert.AreEqual("<TestConfiguration>Value</TestConfiguration>", handler.Configuration.OuterXml);

            handler = template.Providers.FirstOrDefault(p => p.HandlerType == "MyType2, MyAssembly2");
            Assert.IsNotNull(handler);
            Assert.IsFalse(handler.Enabled);
            Assert.IsNull(handler.Configuration);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_AuditSettings_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);

            Assert.IsNotNull(template.AuditSettings);
            Assert.IsTrue(template.AuditSettings.AuditFlags.Has(AuditMaskType.CheckIn));
            Assert.IsTrue(template.AuditSettings.AuditFlags.Has(AuditMaskType.CheckOut));
            Assert.IsTrue(template.AuditSettings.AuditFlags.Has(AuditMaskType.Search));
            Assert.AreEqual(10, template.AuditSettings.AuditLogTrimmingRetention);
            Assert.IsTrue(template.AuditSettings.TrimAuditLog);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_AuditSettings_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.AuditSettings = new AuditSettings() {
                AuditFlags = AuditMaskType.ProfileChange | AuditMaskType.Move,
                AuditLogTrimmingRetention = 10,
                TrimAuditLog = true

            };

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-audit.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-audit.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.AuditSettings);
            Assert.IsNotNull(template.AuditSettings.Audit);
            Assert.AreEqual(2, template.AuditSettings.Audit.Length);
            Assert.IsNotNull(template.AuditSettings.Audit.FirstOrDefault(a=> a.AuditFlag == Core.Framework.Provisioning.Providers.Xml.V201605.AuditSettingsAuditAuditFlag.ProfileChange));
            Assert.IsNotNull(template.AuditSettings.Audit.FirstOrDefault(a => a.AuditFlag == Core.Framework.Provisioning.Providers.Xml.V201605.AuditSettingsAuditAuditFlag.Move));
            Assert.AreEqual(10, template.AuditSettings.AuditLogTrimmingRetention);
            Assert.IsTrue(template.AuditSettings.AuditLogTrimmingRetentionSpecified);
            Assert.IsTrue(template.AuditSettings.TrimAuditLog);
            Assert.IsTrue(template.AuditSettings.TrimAuditLogSpecified);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Features_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.Features);
            Assert.IsNotNull(template.Features.SiteFeatures);
            Assert.AreEqual(3, template.Features.SiteFeatures.Count);
            var feature = template.Features.SiteFeatures.FirstOrDefault(f => f.Id == new Guid("b50e3104-6812-424f-a011-cc90e6327318"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.SiteFeatures.FirstOrDefault(f => f.Id == new Guid("9c0834e1-ba47-4d49-812b-7d4fb6fea211"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.SiteFeatures.FirstOrDefault(f => f.Id == new Guid("0af5989a-3aea-4519-8ab0-85d91abe39ff"));
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);

            Assert.IsNotNull(template.Features.WebFeatures);
            Assert.AreEqual(4, template.Features.WebFeatures.Count);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("7201d6a4-a5d3-49a1-8c19-19c4bac6e668"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("961d6a9c-4388-4cf2-9733-38ee8c89afd4"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("e2f2bb18-891d-4812-97df-c265afdba297"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("4aec7207-0d02-4f4f-aa07-b370199cd0c7"));
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Features_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.Features = new Features();

            result.Features.SiteFeatures.Add(new Core.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("d8f187e3-2bf3-43a3-99a0-024edaffab5e")
            });
            result.Features.SiteFeatures.Add(new Core.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("89c029c5-d289-4936-8ba6-6f3386a8a03f"),
                Deactivate = true
            });
            result.Features.WebFeatures.Add(new Core.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("a22d7848-6d17-47b5-9c1c-cecc98a6b258")
            });
            result.Features.WebFeatures.Add(new Core.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("d60aed53-05f3-4d1c-a12f-677da19a8c31"),
                Deactivate = true
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-features.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-features.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.Features);
            Assert.IsNotNull(template.Features.SiteFeatures);
            Assert.AreEqual(2, template.Features.SiteFeatures.Length);
            var feature = template.Features.SiteFeatures.FirstOrDefault(f => f.ID == "d8f187e3-2bf3-43a3-99a0-024edaffab5e");
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.SiteFeatures.FirstOrDefault(f => f.ID == "89c029c5-d289-4936-8ba6-6f3386a8a03f");
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);

            Assert.IsNotNull(template.Features.WebFeatures);
            Assert.AreEqual(2, template.Features.WebFeatures.Length);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.ID == "a22d7848-6d17-47b5-9c1c-cecc98a6b258");
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.ID == "d60aed53-05f3-4d1c-a12f-677da19a8c31");
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_LocalizationSettings_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            var localizations = template.Localizations;
            Assert.AreEqual(2, localizations.Count);
            var locale = localizations.FirstOrDefault(l => l.LCID == 1033);
            Assert.IsNotNull(locale);
            Assert.AreEqual("en-US", locale.Name);
            Assert.AreEqual("template.en-US.resx", locale.ResourceFile);

            locale = localizations.FirstOrDefault(l => l.LCID == 1040);
            Assert.IsNotNull(locale);
            Assert.AreEqual("it-IT", locale.Name);
            Assert.AreEqual("template.it-It.resx", locale.ResourceFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_LocalizationSettings_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.Localizations.Add(new Localization()
            {
                LCID = 1033,
                Name = "en-US",
                ResourceFile = "template.en-US.resx"
            });
            result.Localizations.Add(new Localization()
            {
                LCID = 1040,
                Name = "it-IT",
                ResourceFile = "template.it-It.resx"
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-local.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-local.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            var localizations = wrappedResult.Localizations;
            Assert.AreEqual(2, localizations.Length);
            var locale = localizations.FirstOrDefault(l => l.LCID == 1033);
            Assert.IsNotNull(locale);
            Assert.AreEqual("en-US", locale.Name);
            Assert.AreEqual("template.en-US.resx", locale.ResourceFile);

            locale = localizations.FirstOrDefault(l => l.LCID == 1040);
            Assert.IsNotNull(locale);
            Assert.AreEqual("it-IT", locale.Name);
            Assert.AreEqual("template.it-It.resx", locale.ResourceFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_WebSettings_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.WebSettings);
            Assert.AreEqual("Resources/Themes/Contoso/Contoso.css", template.WebSettings.AlternateCSS);
            Assert.AreEqual("seattle.master", template.WebSettings.MasterPageUrl);
            Assert.AreEqual("custom.master", template.WebSettings.CustomMasterPageUrl);
            Assert.IsTrue(template.WebSettings.NoCrawl);
            Assert.AreEqual("admin@contoso.com", template.WebSettings.RequestAccessEmail);
            Assert.AreEqual("Resources/Themes/Contoso/contosologo.png", template.WebSettings.SiteLogo);
            Assert.AreEqual("Contoso Portal", template.WebSettings.Title);
            Assert.AreEqual("/Pages/home.aspx", template.WebSettings.WelcomePage);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_WebSettings_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.WebSettings = new WebSettings() {
                AlternateCSS = "Resources/Themes/Contoso/Contoso.css",
                MasterPageUrl= "seattle.master",
                CustomMasterPageUrl="custom.master",
                Description="Test site",
                NoCrawl=true,
                RequestAccessEmail="admin@contoso.com",
                SiteLogo = "Resources/Themes/Contoso/contosologo.png",
                Title="Contoso Portal",
                WelcomePage="/Pages/home.aspx"
            };
            
            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-web.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-web.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.WebSettings);
            Assert.AreEqual("Resources/Themes/Contoso/Contoso.css", template.WebSettings.AlternateCSS);
            Assert.AreEqual("seattle.master", template.WebSettings.MasterPageUrl);
            Assert.AreEqual("custom.master", template.WebSettings.CustomMasterPageUrl);
            Assert.IsTrue(template.WebSettings.NoCrawl);
            Assert.AreEqual("admin@contoso.com", template.WebSettings.RequestAccessEmail);
            Assert.AreEqual("Resources/Themes/Contoso/contosologo.png", template.WebSettings.SiteLogo);
            Assert.AreEqual("Contoso Portal", template.WebSettings.Title);
            Assert.AreEqual("/Pages/home.aspx", template.WebSettings.WelcomePage);
            Assert.AreEqual("Test site", template.WebSettings.Description);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_RegionalSettings_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.RegionalSettings);
            Assert.AreEqual(2, template.RegionalSettings.AdjustHijriDays);
            Assert.AreEqual(CalendarType.GregorianArabic, template.RegionalSettings.AlternateCalendarType);
            Assert.AreEqual(CalendarType.Gregorian, template.RegionalSettings.CalendarType);
            Assert.AreEqual(1, template.RegionalSettings.Collation);
            Assert.AreEqual(DayOfWeek.Sunday, template.RegionalSettings.FirstDayOfWeek);
            Assert.AreEqual(1, template.RegionalSettings.FirstWeekOfYear);
            Assert.AreEqual(1040, template.RegionalSettings.LocaleId);
            Assert.IsTrue(template.RegionalSettings.ShowWeeks);
            Assert.IsTrue(template.RegionalSettings.Time24);
            Assert.AreEqual(2, template.RegionalSettings.TimeZone);
            Assert.AreEqual(WorkHour.PM0600, template.RegionalSettings.WorkDayEndHour);
            Assert.AreEqual(5, template.RegionalSettings.WorkDays);
            Assert.AreEqual(WorkHour.AM0900, template.RegionalSettings.WorkDayStartHour);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_RegionalSettings_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.RegionalSettings = new Core.Framework.Provisioning.Model.RegionalSettings()
            {
                AdjustHijriDays = 2,
                AlternateCalendarType = CalendarType.GregorianArabic,
                CalendarType = CalendarType.Gregorian,
                Collation = 1,
                FirstDayOfWeek = DayOfWeek.Sunday,
                FirstWeekOfYear = 1,
                LocaleId = 1040,
                ShowWeeks = true,
                Time24 = true,
                TimeZone = 2,
                WorkDayEndHour = WorkHour.PM0600,
                WorkDays = 5,
                WorkDayStartHour = WorkHour.AM0900
            };

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-region.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-region.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.RegionalSettings);
            Assert.AreEqual(2, template.RegionalSettings.AdjustHijriDays);
            Assert.IsTrue(template.RegionalSettings.AdjustHijriDaysSpecified);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.CalendarType.GregorianArabicCalendar, template.RegionalSettings.AlternateCalendarType);
            Assert.IsTrue(template.RegionalSettings.AlternateCalendarTypeSpecified);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.CalendarType.Gregorian, template.RegionalSettings.CalendarType);
            Assert.IsTrue(template.RegionalSettings.CalendarTypeSpecified);
            Assert.AreEqual(1, template.RegionalSettings.Collation);
            Assert.IsTrue(template.RegionalSettings.CollationSpecified);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.DayOfWeek.Sunday, template.RegionalSettings.FirstDayOfWeek);
            Assert.IsTrue(template.RegionalSettings.FirstDayOfWeekSpecified);
            Assert.AreEqual(1, template.RegionalSettings.FirstWeekOfYear);
            Assert.IsTrue(template.RegionalSettings.FirstWeekOfYearSpecified);
            Assert.AreEqual(1040, template.RegionalSettings.LocaleId);
            Assert.IsTrue(template.RegionalSettings.LocaleIdSpecified);
            Assert.IsTrue(template.RegionalSettings.ShowWeeks);
            Assert.IsTrue(template.RegionalSettings.ShowWeeksSpecified);
            Assert.IsTrue(template.RegionalSettings.Time24);
            Assert.IsTrue(template.RegionalSettings.Time24Specified);
            Assert.AreEqual("2", template.RegionalSettings.TimeZone);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.WorkHour.Item600PM, template.RegionalSettings.WorkDayEndHour);
            Assert.IsTrue(template.RegionalSettings.WorkDayEndHourSpecified);
            Assert.AreEqual(5, template.RegionalSettings.WorkDays);
            Assert.IsTrue(template.RegionalSettings.WorkDaysSpecified);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.WorkHour.Item900AM, template.RegionalSettings.WorkDayStartHour);
            Assert.IsTrue(template.RegionalSettings.WorkDayStartHourSpecified);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SupportedUILanguages_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.SupportedUILanguages);
            Assert.AreEqual(2, template.SupportedUILanguages.Count);
            Assert.IsNotNull(template.SupportedUILanguages.FirstOrDefault(l => l.LCID == 1033));
            Assert.IsNotNull(template.SupportedUILanguages.FirstOrDefault(l => l.LCID == 1040));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SupportedUILanguages_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.SupportedUILanguages.Add(new SupportedUILanguage() { LCID = 1040 });
            result.SupportedUILanguages.Add(new SupportedUILanguage() { LCID = 1033 });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-lang.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-lang.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.SupportedUILanguages);
            Assert.AreEqual(2, template.SupportedUILanguages.Length);
            Assert.IsNotNull(template.SupportedUILanguages.FirstOrDefault(l => l.LCID == 1033));
            Assert.IsNotNull(template.SupportedUILanguages.FirstOrDefault(l => l.LCID == 1040));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_PropertyBagEntries_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.PropertyBagEntries);
            Assert.AreEqual(2, template.PropertyBagEntries.Count);
            var prop = template.PropertyBagEntries.FirstOrDefault(p => p.Key == "KEY1");
            Assert.IsNotNull(prop);
            Assert.AreEqual("value1", prop.Value);
            Assert.IsTrue(prop.Indexed);
            Assert.IsTrue(prop.Overwrite);
            prop = template.PropertyBagEntries.FirstOrDefault(p => p.Key == "KEY2");
            Assert.IsNotNull(prop);
            Assert.AreEqual("value2", prop.Value);
            Assert.IsFalse(prop.Indexed);
            Assert.IsFalse(prop.Overwrite);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_PropertyBagEntries_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.PropertyBagEntries.Add(new PropertyBagEntry() {
                Key = "KEY1",
                Value = "value1",
                Overwrite = true,
                Indexed = true
            });
            result.PropertyBagEntries.Add(new PropertyBagEntry()
            {
                Key = "KEY2",
                Value = "value2"
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-propbag.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-propbag.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.PropertyBagEntries);
            Assert.AreEqual(2, template.PropertyBagEntries.Length);
            var prop = template.PropertyBagEntries.FirstOrDefault(p => p.Key == "KEY1");
            Assert.IsNotNull(prop);
            Assert.AreEqual("value1", prop.Value);
            Assert.IsTrue(prop.Indexed);
            Assert.IsTrue(prop.Overwrite);
            Assert.IsTrue(prop.OverwriteSpecified);
            prop = template.PropertyBagEntries.FirstOrDefault(p => p.Key == "KEY2");
            Assert.IsNotNull(prop);
            Assert.AreEqual("value2", prop.Value);
            Assert.IsFalse(prop.Indexed);
            Assert.IsFalse(prop.Overwrite);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TemplateParameters_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            var param = template.Parameters.FirstOrDefault(p => p.Key == "Parameter1");
            Assert.IsNotNull(param);
            Assert.AreEqual("ValueParameter1", param.Value);
            param = template.Parameters.FirstOrDefault(p => p.Key == "Parameter2");
            Assert.IsNotNull(param);
            Assert.AreEqual("ValueParameter2", param.Value);
            param = template.Parameters.FirstOrDefault(p => p.Key == "Parameter3");
            Assert.IsNotNull(param);
            Assert.AreEqual("ValueParameter3", param.Value);
            param = template.Parameters.FirstOrDefault(p => p.Key == "Parameter4");
            Assert.IsNotNull(param);
            Assert.AreEqual("ValueParameter4", param.Value);
            param = template.Parameters.FirstOrDefault(p => p.Key == "Parameter5");
            Assert.IsNotNull(param);
            Assert.AreEqual("ValueParameter5", param.Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TemplateParameters_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();

            result.Parameters.Add("Parameter1", "ValueParameter1");
            result.Parameters.Add("Parameter2", "ValueParameter2");
            result.Parameters.Add("Parameter3", "ValueParameter3");
            result.Parameters.Add("Parameter4", "ValueParameter4");
            result.Parameters.Add("Parameter5", "ValueParameter5");

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-temppar.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-temppar.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Preferences);
            Assert.IsNotNull(wrappedResult.Preferences.Parameters);
            Assert.AreEqual(5, wrappedResult.Preferences.Parameters.Length);

            var param = wrappedResult.Preferences.Parameters.FirstOrDefault(p => p.Key == "Parameter1");
            Assert.IsNotNull(param);
            Assert.IsNotNull(param.Text);
            Assert.AreEqual(1, param.Text.Length);
            Assert.AreEqual("ValueParameter1", param.Text.First());
            param = wrappedResult.Preferences.Parameters.FirstOrDefault(p => p.Key == "Parameter2");
            Assert.IsNotNull(param);
            Assert.IsNotNull(param.Text);
            Assert.AreEqual(1, param.Text.Length);
            Assert.AreEqual("ValueParameter2", param.Text.First());
            param = wrappedResult.Preferences.Parameters.FirstOrDefault(p => p.Key == "Parameter3");
            Assert.IsNotNull(param);
            Assert.IsNotNull(param.Text);
            Assert.AreEqual(1, param.Text.Length);
            Assert.AreEqual("ValueParameter3", param.Text.First());
            param = wrappedResult.Preferences.Parameters.FirstOrDefault(p => p.Key == "Parameter4");
            Assert.IsNotNull(param);
            Assert.IsNotNull(param.Text);
            Assert.AreEqual(1, param.Text.Length);
            Assert.AreEqual("ValueParameter4", param.Text.First());
            param = wrappedResult.Preferences.Parameters.FirstOrDefault(p => p.Key == "Parameter5");
            Assert.IsNotNull(param);
            Assert.IsNotNull(param.Text);
            Assert.AreEqual(1, param.Text.Length);
            Assert.AreEqual("ValueParameter5", param.Text.First());
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TemplateBaseProperties_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.AreEqual(1.2, template.Version);
            var param = template.Properties.FirstOrDefault(p => p.Key == "Key1");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value1", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key2");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value2", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key3");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value3", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key4");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value4", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key5");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value5", param.Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TemplateBaseProperties_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();
            result.Version = 1.2;
            result.Properties.Add("Key1", "Value1");
            result.Properties.Add("Key2", "Value2");
            result.Properties.Add("Key3", "Value3");
            result.Properties.Add("Key4", "Value4");
            result.Properties.Add("Key5", "Value5");

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-tempprop.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-tempprop.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.AreEqual((decimal)1.2, template.Version);

            Assert.IsNotNull(template.Properties);
            Assert.AreEqual(5, template.Properties.Length);
            var param = template.Properties.FirstOrDefault(p => p.Key == "Key1");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value1", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key2");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value2", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key3");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value3", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key4");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value4", param.Value);
            param = template.Properties.FirstOrDefault(p => p.Key == "Key5");
            Assert.IsNotNull(param);
            Assert.AreEqual("Value5", param.Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SiteFields_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.SiteFields);
            Assert.AreEqual(4, template.SiteFields.Count);
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"Project ID\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" Required=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"Project Name\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{A5DE9600-B7A6-42DD-A05E-10D4F1500208}\" Type=\"Text\" Name=\"ProjectManager\" DisplayName=\"Project Manager\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{F1A1715E-6C52-40DE-8403-E9AAFD0470D0}\" Type=\"Text\" Name=\"DocumentDescription\" DisplayName=\"Document Description\" Group=\"My Columns \" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SiteFields_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();
            result.SiteFields.Add(new Core.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"Project ID\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" Required=\"TRUE\" />"
            });
            result.SiteFields.Add(new Core.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID = \"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"Project Name\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });
            result.SiteFields.Add(new Core.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID = \"{A5DE9600-B7A6-42DD-A05E-10D4F1500208}\" Type=\"Text\" Name=\"ProjectManager\" DisplayName=\"Project Manager\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });
            result.SiteFields.Add(new Core.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID = \"{F1A1715E-6C52-40DE-8403-E9AAFD0470D0}\" Type=\"Text\" Name=\"DocumentDescription\" DisplayName=\"Document Description\" Group=\"My Columns \" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-flds.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-flds.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.SiteFields);
            Assert.IsNotNull(template.SiteFields.Any);
            Assert.AreEqual(4, template.SiteFields.Any.Length);
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"Project ID\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" Required=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"Project Name\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{A5DE9600-B7A6-42DD-A05E-10D4F1500208}\" Type=\"Text\" Name=\"ProjectManager\" DisplayName=\"Project Manager\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{F1A1715E-6C52-40DE-8403-E9AAFD0470D0}\" Type=\"Text\" Name=\"DocumentDescription\" DisplayName=\"Document Description\" Group=\"My Columns \" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Navigation_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.Navigation);
            Assert.IsNotNull(template.Navigation.GlobalNavigation);
            Assert.AreEqual(GlobalNavigationType.Managed, template.Navigation.GlobalNavigation.NavigationType);
            Assert.IsNull(template.Navigation.GlobalNavigation.StructuralNavigation);
            Assert.IsNotNull(template.Navigation.GlobalNavigation.ManagedNavigation);
            Assert.AreEqual("415185a1-ee1c-4ce9-9e38-cea3f854e802", template.Navigation.GlobalNavigation.ManagedNavigation.TermSetId);
            Assert.AreEqual("c1175ad1-c710-4131-a6c9-aa854a5cc4c4", template.Navigation.GlobalNavigation.ManagedNavigation.TermStoreId);

            Assert.IsNotNull(template.Navigation.CurrentNavigation);
            Assert.AreEqual(CurrentNavigationType.Structural, template.Navigation.CurrentNavigation.NavigationType);
            Assert.IsNull(template.Navigation.CurrentNavigation.ManagedNavigation);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation);
            Assert.IsTrue(template.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes);
            Assert.AreEqual(2, template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.Count);

            var n1 = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.FirstOrDefault(n => n.Title == "Node 1");
            Assert.IsNotNull(n1);
            Assert.AreEqual("/Node1.aspx", n1.Url);
            Assert.IsFalse(n1.IsExternal);
            Assert.IsNotNull(n1.NavigationNodes);
            Assert.AreEqual(2, n1.NavigationNodes.Count);

            var n11 = n1.NavigationNodes.FirstOrDefault(n => n.Title == "Node 1.1");
            Assert.IsNotNull(n11);
            Assert.AreEqual("http://aka.ms/SharePointPnP", n11.Url);
            Assert.IsTrue(n11.IsExternal);
            Assert.IsNotNull(n11.NavigationNodes);
            Assert.AreEqual(1, n11.NavigationNodes.Count);

            var n111 = n11.NavigationNodes.FirstOrDefault(n => n.Title == "Node 1.1.1");
            Assert.IsNotNull(n111);
            Assert.AreEqual("http://aka.ms/OfficeDevPnP", n111.Url);
            Assert.IsTrue(n111.IsExternal);
            Assert.IsNotNull(n111.NavigationNodes);
            Assert.AreEqual(0, n111.NavigationNodes.Count);

            var n12 = n1.NavigationNodes.FirstOrDefault(n => n.Title == "Node 1.2");
            Assert.IsNotNull(n12);
            Assert.AreEqual("/Node1-2.aspx", n12.Url);
            Assert.IsTrue(n12.IsExternal);
            Assert.IsNotNull(n12.NavigationNodes);
            Assert.AreEqual(0, n12.NavigationNodes.Count);

            var n2 = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.FirstOrDefault(n => n.Title == "Node 2");
            Assert.IsNotNull(n2);
            Assert.AreEqual("/Node1.aspx", n2.Url);
            Assert.IsFalse(n2.IsExternal);
            Assert.IsNotNull(n2.NavigationNodes);
            Assert.AreEqual(0, n2.NavigationNodes.Count);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Navigation_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();
            result.Navigation = new Core.Framework.Provisioning.Model.Navigation(
                new GlobalNavigation(GlobalNavigationType.Managed, null, new ManagedNavigation()), 
                new CurrentNavigation(CurrentNavigationType.Structural, new StructuralNavigation(), null));
            
            result.Navigation.GlobalNavigation.ManagedNavigation.TermSetId = "415185a1-ee1c-4ce9-9e38-cea3f854e802";
            result.Navigation.GlobalNavigation.ManagedNavigation.TermStoreId = "c1175ad1-c710-4131-a6c9-aa854a5cc4c4";

            result.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes = true;
            var node1 = new Core.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = false,
                Title = "Node 1",
                Url = "/Node1.aspx",
                
            };
            var node11 = new Core.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = true,
                Title = "Node 1.1",
                Url = "http://aka.ms/SharePointPnP"
            };
            var node111 = new Core.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = true,
                Title = "Node 1.1.1",
                Url = "http://aka.ms/OfficeDevPnP"
            };
            var node12 = new Core.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = true,
                Title = "Node 1.2",
                Url = "/Node1-2.aspx"
            };
            var node2 = new Core.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = false,
                Title = "Node 2",
                Url = "/Node1.aspx"
            };
            node11.NavigationNodes.Add(node111);
            node1.NavigationNodes.Add(node11);
            node1.NavigationNodes.Add(node12);
            result.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.Add(node1);
            result.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.Add(node2);

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-nav.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-nav.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.Navigation);
            Assert.IsNotNull(template.Navigation.GlobalNavigation);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.NavigationGlobalNavigationNavigationType.Managed, template.Navigation.GlobalNavigation.NavigationType);
            Assert.IsNull(template.Navigation.GlobalNavigation.StructuralNavigation);
            Assert.IsNotNull(template.Navigation.GlobalNavigation.ManagedNavigation);
            Assert.AreEqual("415185a1-ee1c-4ce9-9e38-cea3f854e802", template.Navigation.GlobalNavigation.ManagedNavigation.TermSetId);
            Assert.AreEqual("c1175ad1-c710-4131-a6c9-aa854a5cc4c4", template.Navigation.GlobalNavigation.ManagedNavigation.TermStoreId);

            Assert.IsNotNull(template.Navigation.CurrentNavigation);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.NavigationCurrentNavigationNavigationType.Structural, template.Navigation.CurrentNavigation.NavigationType);
            Assert.IsNull(template.Navigation.CurrentNavigation.ManagedNavigation);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation);
            Assert.IsTrue(template.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode);
            Assert.AreEqual(2, template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode.Length);

            var n1 = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode.FirstOrDefault(n => n.Title == "Node 1");
            Assert.IsNotNull(n1);
            Assert.AreEqual("/Node1.aspx", n1.Url);
            Assert.IsFalse(n1.IsExternal);
            Assert.IsNotNull(n1.NavigationNode1);
            Assert.AreEqual(2, n1.NavigationNode1.Length);

            var n11 = n1.NavigationNode1.FirstOrDefault(n => n.Title == "Node 1.1");
            Assert.IsNotNull(n11);
            Assert.AreEqual("http://aka.ms/SharePointPnP", n11.Url);
            Assert.IsTrue(n11.IsExternal);
            Assert.IsNotNull(n11.NavigationNode1);
            Assert.AreEqual(1, n11.NavigationNode1.Length);

            var n111 = n11.NavigationNode1.FirstOrDefault(n => n.Title == "Node 1.1.1");
            Assert.IsNotNull(n111);
            Assert.AreEqual("http://aka.ms/OfficeDevPnP", n111.Url);
            Assert.IsTrue(n111.IsExternal);
            Assert.IsNull(n111.NavigationNode1);

            var n12 = n1.NavigationNode1.FirstOrDefault(n => n.Title == "Node 1.2");
            Assert.IsNotNull(n12);
            Assert.AreEqual("/Node1-2.aspx", n12.Url);
            Assert.IsTrue(n12.IsExternal);
            Assert.IsNull(n12.NavigationNode1);

            var n2 = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode.FirstOrDefault(n => n.Title == "Node 2");
            Assert.IsNotNull(n2);
            Assert.AreEqual("/Node1.aspx", n2.Url);
            Assert.IsFalse(n2.IsExternal);
            Assert.IsNull(n2.NavigationNode1);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Security_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.Security);
            Assert.IsTrue(template.Security.BreakRoleInheritance);
            Assert.IsTrue(template.Security.ClearSubscopes);
            Assert.IsTrue(template.Security.CopyRoleAssignments);

            Assert.IsNotNull(template.Security.AdditionalAdministrators);
            Assert.AreEqual(2, template.Security.AdditionalAdministrators.Count);
            Assert.IsNotNull(template.Security.AdditionalAdministrators.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalAdministrators.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsNotNull(template.Security.AdditionalOwners);
            Assert.AreEqual(2, template.Security.AdditionalOwners.Count);
            Assert.IsNotNull(template.Security.AdditionalOwners.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalOwners.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsNotNull(template.Security.AdditionalMembers);
            Assert.AreEqual(2, template.Security.AdditionalMembers.Count);
            Assert.IsNotNull(template.Security.AdditionalMembers.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalMembers.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsNotNull(template.Security.AdditionalVisitors);
            Assert.AreEqual(2, template.Security.AdditionalVisitors.Count);
            Assert.IsNotNull(template.Security.AdditionalVisitors.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalVisitors.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));

            Assert.IsNotNull(template.Security.SiteSecurityPermissions);
            Assert.IsNotNull(template.Security.SiteSecurityPermissions.RoleDefinitions);
            Assert.AreEqual(2, template.Security.SiteSecurityPermissions.RoleDefinitions.Count);
            var role = template.Security.SiteSecurityPermissions.RoleDefinitions.FirstOrDefault(r => r.Name == "User");
            Assert.IsNotNull(role);
            Assert.AreEqual("User Role", role.Description);
            Assert.IsNotNull(role.Permissions);
            Assert.AreEqual(2, role.Permissions.Count);
            Assert.IsTrue(role.Permissions.Contains(PermissionKind.ViewListItems));
            Assert.IsTrue(role.Permissions.Contains(PermissionKind.AddListItems));

            role = template.Security.SiteSecurityPermissions.RoleDefinitions.FirstOrDefault(r => r.Name == "EmptyRole");
            Assert.IsNotNull(role);
            Assert.AreEqual("Empty Role", role.Description);
            Assert.IsNotNull(role.Permissions);
            Assert.AreEqual(1, role.Permissions.Count);
            Assert.IsTrue(role.Permissions.Contains(PermissionKind.EmptyMask));

            Assert.IsNotNull(template.Security.SiteSecurityPermissions.RoleAssignments);
            Assert.AreEqual(2, template.Security.SiteSecurityPermissions.RoleAssignments.Count);
            var assign = template.Security.SiteSecurityPermissions.RoleAssignments.FirstOrDefault(p => p.Principal == "admin@contoso.com");
            Assert.IsNotNull(assign);
            Assert.AreEqual("Owner", assign.RoleDefinition);
            assign = template.Security.SiteSecurityPermissions.RoleAssignments.FirstOrDefault(p => p.Principal == "user@contoso.com");
            Assert.IsNotNull(assign);
            Assert.AreEqual("User", assign.RoleDefinition);

            Assert.IsNotNull(template.Security.SiteGroups);
            Assert.AreEqual(2, template.Security.SiteGroups.Count);
            var group = template.Security.SiteGroups.FirstOrDefault(g => g.Title == "TestGroup1");
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Group 1", group.Description);
            Assert.AreEqual("user1@contoso.com", group.Owner);
            Assert.AreEqual("group1@contoso.com", group.RequestToJoinLeaveEmailSetting);
            Assert.IsTrue(group.AllowMembersEditMembership);
            Assert.IsTrue(group.AllowRequestToJoinLeave);
            Assert.IsTrue(group.AutoAcceptRequestToJoinLeave);
            Assert.IsTrue(group.OnlyAllowMembersViewMembership);
            Assert.IsNotNull(group.Members);
            Assert.AreEqual(2, group.Members.Count);
            Assert.IsNotNull(group.Members.FirstOrDefault(m => m.Name == "user1@contoso.com"));
            Assert.IsNotNull(group.Members.FirstOrDefault(m => m.Name == "user2@contoso.com"));

            group = template.Security.SiteGroups.FirstOrDefault(g => g.Title == "TestGroup2");
            Assert.IsNotNull(group);
            Assert.AreEqual("user2@contoso.com", group.Owner);
            Assert.IsTrue(string.IsNullOrEmpty(group.Description));
            Assert.IsTrue(string.IsNullOrEmpty(group.RequestToJoinLeaveEmailSetting));
            Assert.IsFalse(group.AllowMembersEditMembership);
            Assert.IsFalse(group.AllowRequestToJoinLeave);
            Assert.IsFalse(group.AutoAcceptRequestToJoinLeave);
            Assert.IsFalse(group.OnlyAllowMembersViewMembership);
            Assert.IsTrue(group.Members == null || group.Members.Count == 0);

        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Security_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();
            result.Security = new SiteSecurity()
            {
                BreakRoleInheritance = true,
                ClearSubscopes = true,
                CopyRoleAssignments = true
            };
            result.Security.AdditionalAdministrators.Add(new Core.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalAdministrators.Add(new Core.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });
            result.Security.AdditionalOwners.Add(new Core.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalOwners.Add(new Core.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });
            result.Security.AdditionalMembers.Add(new Core.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalMembers.Add(new Core.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });
            result.Security.AdditionalVisitors.Add(new Core.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalVisitors.Add(new Core.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });

            result.Security.SiteSecurityPermissions.RoleDefinitions.Add(new Core.Framework.Provisioning.Model.RoleDefinition(new List<PermissionKind>() {
                PermissionKind.ViewListItems,
                PermissionKind.AddListItems
            })
            {
                Name = "User",
                Description = "User Role"
            });
            result.Security.SiteSecurityPermissions.RoleDefinitions.Add(new Core.Framework.Provisioning.Model.RoleDefinition(new List<PermissionKind>() {
                PermissionKind.EmptyMask
            })
            {
                Name = "EmptyRole",
                Description = "Empty Role"
            });
            result.Security.SiteSecurityPermissions.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment() {
                Principal = "admin@contoso.com",
                RoleDefinition = "Owner"
            });
            result.Security.SiteSecurityPermissions.RoleAssignments.Add(new Core.Framework.Provisioning.Model.RoleAssignment()
            {
                Principal = "user@contoso.com",
                RoleDefinition = "User"
            });

            result.Security.SiteGroups.Add(new SiteGroup(new List<Core.Framework.Provisioning.Model.User>()
            {
                new Core.Framework.Provisioning.Model.User()
                {
                     Name = "user1@contoso.com"
                },
                new Core.Framework.Provisioning.Model.User()
                {
                     Name = "user2@contoso.com"
                }
            })
            {
                AllowMembersEditMembership = true,
                AllowRequestToJoinLeave = true,
                AutoAcceptRequestToJoinLeave = true,
                Description = "Test Group 1",
                OnlyAllowMembersViewMembership = true,
                Owner = "user1@contoso.com",
                RequestToJoinLeaveEmailSetting = "group1@contoso.com",
                Title = "TestGroup1"
            });
            result.Security.SiteGroups.Add(new SiteGroup()
            {
                Title = "TestGroup2",
                Owner = "user2@contoso.com"
            });

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-sec.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-sec.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.Security);
            Assert.IsTrue(template.Security.BreakRoleInheritance);
            Assert.IsTrue(template.Security.ClearSubscopes);
            Assert.IsTrue(template.Security.CopyRoleAssignments);

            Assert.IsNotNull(template.Security.AdditionalAdministrators);
            Assert.AreEqual(2, template.Security.AdditionalAdministrators.Length);
            Assert.IsNotNull(template.Security.AdditionalAdministrators.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalAdministrators.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsNotNull(template.Security.AdditionalOwners);
            Assert.AreEqual(2, template.Security.AdditionalOwners.Length);
            Assert.IsNotNull(template.Security.AdditionalOwners.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalOwners.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsNotNull(template.Security.AdditionalMembers);
            Assert.AreEqual(2, template.Security.AdditionalMembers.Length);
            Assert.IsNotNull(template.Security.AdditionalMembers.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalMembers.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsNotNull(template.Security.AdditionalVisitors);
            Assert.AreEqual(2, template.Security.AdditionalVisitors.Length);
            Assert.IsNotNull(template.Security.AdditionalVisitors.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalVisitors.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));

            Assert.IsNotNull(template.Security.Permissions);
            Assert.IsNotNull(template.Security.Permissions.RoleDefinitions);
            Assert.AreEqual(2, template.Security.Permissions.RoleDefinitions.Length);
            var role = template.Security.Permissions.RoleDefinitions.FirstOrDefault(r => r.Name == "User");
            Assert.IsNotNull(role);
            Assert.AreEqual("User Role", role.Description);
            Assert.IsNotNull(role.Permissions);
            Assert.AreEqual(2, role.Permissions.Length);
            Assert.IsTrue(role.Permissions.Contains(Core.Framework.Provisioning.Providers.Xml.V201605.RoleDefinitionPermission.ViewListItems));
            Assert.IsTrue(role.Permissions.Contains(Core.Framework.Provisioning.Providers.Xml.V201605.RoleDefinitionPermission.AddListItems));

            role = template.Security.Permissions.RoleDefinitions.FirstOrDefault(r => r.Name == "EmptyRole");
            Assert.IsNotNull(role);
            Assert.AreEqual("Empty Role", role.Description);
            Assert.IsNotNull(role.Permissions);
            Assert.AreEqual(1, role.Permissions.Length);
            Assert.IsTrue(role.Permissions.Contains(Core.Framework.Provisioning.Providers.Xml.V201605.RoleDefinitionPermission.EmptyMask));

            Assert.IsNotNull(template.Security.Permissions);
            Assert.IsNotNull(template.Security.Permissions.RoleAssignments);
            Assert.AreEqual(2, template.Security.Permissions.RoleAssignments.Length);
            var assign = template.Security.Permissions.RoleAssignments.FirstOrDefault(p => p.Principal == "admin@contoso.com");
            Assert.IsNotNull(assign);
            Assert.AreEqual("Owner", assign.RoleDefinition);
            assign = template.Security.Permissions.RoleAssignments.FirstOrDefault(p => p.Principal == "user@contoso.com");
            Assert.IsNotNull(assign);
            Assert.AreEqual("User", assign.RoleDefinition);

            Assert.IsNotNull(template.Security.SiteGroups);
            Assert.AreEqual(2, template.Security.SiteGroups.Length);
            var group = template.Security.SiteGroups.FirstOrDefault(g => g.Title == "TestGroup1");
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Group 1", group.Description);
            Assert.AreEqual("user1@contoso.com", group.Owner);
            Assert.AreEqual("group1@contoso.com", group.RequestToJoinLeaveEmailSetting);
            Assert.IsTrue(group.AllowMembersEditMembership);
            Assert.IsTrue(group.AllowMembersEditMembershipSpecified);
            Assert.IsTrue(group.AllowRequestToJoinLeave);
            Assert.IsTrue(group.AllowRequestToJoinLeaveSpecified);
            Assert.IsTrue(group.AutoAcceptRequestToJoinLeave);
            Assert.IsTrue(group.AutoAcceptRequestToJoinLeaveSpecified);
            Assert.IsTrue(group.OnlyAllowMembersViewMembership);
            Assert.IsTrue(group.OnlyAllowMembersViewMembershipSpecified);
            Assert.IsNotNull(group.Members);
            Assert.AreEqual(2, group.Members.Length);
            Assert.IsNotNull(group.Members.FirstOrDefault(m => m.Name == "user1@contoso.com"));
            Assert.IsNotNull(group.Members.FirstOrDefault(m => m.Name == "user2@contoso.com"));

            group = template.Security.SiteGroups.FirstOrDefault(g => g.Title == "TestGroup2");
            Assert.IsNotNull(group);
            Assert.AreEqual("user2@contoso.com", group.Owner);
            Assert.IsTrue(string.IsNullOrEmpty(group.Description));
            Assert.IsTrue(string.IsNullOrEmpty(group.RequestToJoinLeaveEmailSetting));
            Assert.IsFalse(group.AllowMembersEditMembership);
            Assert.IsFalse(group.AllowRequestToJoinLeave);
            Assert.IsFalse(group.AutoAcceptRequestToJoinLeave);
            Assert.IsFalse(group.OnlyAllowMembersViewMembership);
            Assert.IsNull(group.Members);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ListInstances_201605()
        {
            XMLTemplateProvider provider =
               new XMLFileSystemTemplateProvider(
                   String.Format(@"{0}\..\..\Resources",
                   AppDomain.CurrentDomain.BaseDirectory),
                   "Templates");

            var serializer = new XMLPnPSchemaV201605Serializer();
            var template = provider.GetTemplate("ProvisioningTemplate-2016-05-Sample-03.xml", serializer);
            Assert.IsNotNull(template.Lists);
            Assert.AreEqual(1, template.Lists.Count);

            var l = template.Lists.FirstOrDefault(ls => ls.Title == "Project Documents");
            Assert.IsNotNull(l);
            Assert.IsTrue(l.ContentTypesEnabled);
            Assert.AreEqual("Project Documents are stored here", l.Description);
            Assert.AreEqual("document.dotx", l.DocumentTemplate);
            Assert.AreEqual(1, l.DraftVersionVisibility);
            Assert.IsTrue(l.EnableAttachments);
            Assert.IsTrue(l.EnableFolderCreation);
            Assert.IsTrue(l.EnableMinorVersions);
            Assert.IsTrue(l.EnableModeration);
            Assert.IsTrue(l.EnableVersioning);
            Assert.IsTrue(l.ForceCheckout);
            Assert.IsTrue(l.Hidden);
            Assert.AreEqual(10, l.MaxVersionLimit);
            Assert.AreEqual(2, l.MinorVersionLimit);
            Assert.IsTrue(l.OnQuickLaunch);
            Assert.IsTrue(l.RemoveExistingContentTypes);
            Assert.AreEqual(new Guid("30FB193E-016E-45A6-B6FD-C6C2B31AA150"), l.TemplateFeatureID);
            Assert.AreEqual(101, l.TemplateType);
            Assert.AreEqual("Lists/ProjectDocuments", l.Url);

            var security = l.Security;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(3, security.RoleAssignments.Count);
            var ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            Assert.IsNotNull(l.ContentTypeBindings);
            Assert.AreEqual(3, l.ContentTypeBindings.Count);
            var ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeId == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.IsNotNull(ct);
            Assert.IsTrue(ct.Default);
            Assert.IsFalse(ct.Remove);
            ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeId == "0x0101");
            Assert.IsNotNull(ct);
            Assert.IsFalse(ct.Default);
            Assert.IsTrue(ct.Remove);
            ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeId == "0x0102");
            Assert.IsNotNull(ct);
            Assert.IsFalse(ct.Default);
            Assert.IsFalse(ct.Remove);

            Assert.IsNotNull(l.FieldDefaults);
            Assert.AreEqual(4, l.FieldDefaults.Count);
            var fd = l.FieldDefaults.FirstOrDefault(f => f.Key == "Field01");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue01", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.Key == "Field02");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue02", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.Key == "Field03");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue03", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.Key == "Field04");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue04", fd.Value);

            Assert.IsNotNull(l.DataRows);
            Assert.AreEqual(3, l.DataRows.Count);
#region data row 1 asserts
            var dr = l.DataRows.FirstOrDefault(r => r.Values.Any(d => d.Value.StartsWith("Value01")));
            Assert.IsNotNull(dr);
            security = dr.Security;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(3, security.RoleAssignments.Count);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            var dv = dr.Values.FirstOrDefault(d => d.Key == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-01", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-02", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-03", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-04", dv.Value);
#endregion
#region data row 2 asserts
            dr = l.DataRows.FirstOrDefault(r => r.Values.Any(d => d.Value.StartsWith("Value02")));
            Assert.IsNotNull(dr);
            security = dr.Security;
            Assert.IsNotNull(security);
            Assert.IsFalse(security.ClearSubscopes);
            Assert.IsFalse(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(3, security.RoleAssignments.Count);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            dv = dr.Values.FirstOrDefault(d => d.Key == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-01", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-02", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-03", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-04", dv.Value);
#endregion
#region data row 3 asserts
            dr = l.DataRows.FirstOrDefault(r => r.Values.Any(d => d.Value.StartsWith("Value03")));
            Assert.IsNotNull(dr);
            Assert.IsTrue(dr.Security == null || dr.Security.RoleAssignments == null || dr.Security.RoleAssignments.Count == 0);

            dv = dr.Values.FirstOrDefault(d => d.Key == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-01", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-02", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-03", dv.Value);
            dv = dr.Values.FirstOrDefault(d => d.Key == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-04", dv.Value);
#endregion

#region user custom action
            Assert.IsNotNull(l.UserCustomActions);
            Assert.AreEqual(1, l.UserCustomActions.Count);
            var ua = l.UserCustomActions.FirstOrDefault(a => a.Name == "SampleCustomAction");
            Assert.IsNotNull(ua);
            Assert.AreEqual("Just a sample custom action", ua.Description);
            Assert.IsTrue(ua.Enabled);
            Assert.AreEqual("Samples", ua.Group);
            Assert.AreEqual("OneImage.png", ua.ImageUrl);
            Assert.AreEqual("Any", ua.Location);
            Assert.AreEqual("0x0101", ua.RegistrationId);
            Assert.AreEqual(UserCustomActionRegistrationType.ContentType, ua.RegistrationType);
            Assert.AreEqual(100, ua.Sequence);
            Assert.AreEqual("scriptblock", ua.ScriptBlock);
            Assert.AreEqual("script.js", ua.ScriptSrc);
            Assert.AreEqual("http://somewhere.com/", ua.Url);
            Assert.AreEqual("Sample Action", ua.Title);
            Assert.IsTrue(ua.Remove);
            Assert.IsNotNull(ua.CommandUIExtension);
            Assert.AreEqual(1, ua.CommandUIExtension.Nodes().Count());
            Assert.IsNotNull(ua.Rights);
            Assert.IsTrue(ua.Rights.Has(PermissionKind.AddListItems));
#endregion

            Assert.IsNotNull(l.Views);
            Assert.AreEqual(2, l.Views.Count);

#region field refs
            Assert.IsNotNull(l.FieldRefs);
            Assert.AreEqual(3, l.FieldRefs.Count);
            var fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectID");
            Assert.IsNotNull(fr);
            Assert.AreEqual(new Guid("23203E97-3BFE-40CB-AFB4-07AA2B86BF45"), fr.Id);
            Assert.AreEqual("Project ID", fr.DisplayName);
            Assert.IsFalse(fr.Hidden);
            Assert.IsTrue(fr.Required);
            fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectName");
            Assert.IsNotNull(fr);
            Assert.AreEqual(new Guid("B01B3DBC-4630-4ED1-B5BA-321BC7841E3D"), fr.Id);
            Assert.AreEqual("Project Name", fr.DisplayName);
            Assert.IsTrue(fr.Hidden);
            Assert.IsFalse(fr.Required);
            fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectManager");
            Assert.IsNotNull(fr);
            Assert.AreEqual(new Guid("A5DE9600-B7A6-42DD-A05E-10D4F1500208"), fr.Id);
            Assert.AreEqual("Project Manager", fr.DisplayName);
            Assert.IsFalse(fr.Hidden);
            Assert.IsTrue(fr.Required);
#endregion

#region folders
            Assert.IsNotNull(l.Folders);
            Assert.AreEqual(2, l.Folders.Count);
            var fl = l.Folders.FirstOrDefault(f => f.Name == "Folder02");
            Assert.IsNotNull(fl);
            Assert.IsTrue(fl.Folders == null || fl.Folders.Count == 0);
            fl = l.Folders.FirstOrDefault(f => f.Name == "Folder01");
            Assert.IsNotNull(fl);
            Assert.IsNotNull(fl.Folders);
            var fl1 = fl.Folders.FirstOrDefault(f => f.Name == "Folder01.02");
            Assert.IsNotNull(fl1);
            Assert.IsTrue(fl1.Folders == null || fl1.Folders.Count == 0);
            fl1 = fl.Folders.FirstOrDefault(f => f.Name == "Folder01.01");
            Assert.IsTrue(fl1.Folders == null || fl1.Folders.Count == 0);
            security = fl1.Security;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(3, security.RoleAssignments.Count);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);
#endregion

            Assert.IsNotNull(l.Fields);
            Assert.AreEqual(2, l.Fields.Count);
            Assert.IsTrue(l.Fields.All(x => x.SchemaXml.StartsWith("<Field")));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ListInstances_201605()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = new ProvisioningTemplate();
            var list = new ListInstance()
            {
                Title = "Project Documents",
                ContentTypesEnabled = true,
                Description= "Project Documents are stored here",
                DocumentTemplate = "document.dotx",
                DraftVersionVisibility = 1, 
                EnableAttachments = true,
                EnableFolderCreation = true,
                EnableMinorVersions = true,
                EnableModeration = true,
                EnableVersioning = true,
                ForceCheckout = true,
                Hidden = true,
                MaxVersionLimit = 10,
                MinorVersionLimit = 2,
                OnQuickLaunch = true,
                RemoveExistingContentTypes = true,
                RemoveExistingViews = true,
                TemplateFeatureID = new Guid("30FB193E-016E-45A6-B6FD-C6C2B31AA150"),
                TemplateType = 101,
                Url = "/Lists/ProjectDocuments",
                Security = new ObjectSecurity(new List<Core.Framework.Provisioning.Model.RoleAssignment>()
                {
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal01",
                        RoleDefinition ="Read"
                    },
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal02",
                        RoleDefinition ="Contribute"
                    },
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal03",
                        RoleDefinition ="FullControl"
                    }
                })
                {
                    ClearSubscopes = true,
                    CopyRoleAssignments = true,
                    
                }
            };
            list.ContentTypeBindings.Add(new ContentTypeBinding()
            {
                ContentTypeId = "0x01005D4F34E4BE7F4B6892AEBE088EDD215E",
                Default = true
            });
            list.ContentTypeBindings.Add(new ContentTypeBinding()
            {
                ContentTypeId = "0x0101",
                Remove = true
            });
            list.ContentTypeBindings.Add(new ContentTypeBinding()
            {
                ContentTypeId = "0x0102"
            });

            list.FieldDefaults.Add("Field01", "DefaultValue01");
            list.FieldDefaults.Add("Field02", "DefaultValue02");
            list.FieldDefaults.Add("Field03", "DefaultValue03");
            list.FieldDefaults.Add("Field04", "DefaultValue04");

#region data rows
            list.DataRows.Add(new DataRow(new Dictionary<string, string>() {
                { "Field01", "Value01-01" },
                { "Field02", "Value01-02" },
                { "Field03", "Value01-03" },
                { "Field04", "Value01-04" },
            },
            new ObjectSecurity(new List<Core.Framework.Provisioning.Model.RoleAssignment>() {
                new Core.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal01",
                    RoleDefinition ="Read"
                },
                new Core.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal02",
                    RoleDefinition ="Contribute"
                }
                ,
                new Core.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal03",
                    RoleDefinition ="FullControl"
                }
            })
            {
                ClearSubscopes = true,
                CopyRoleAssignments = true
            }));
            list.DataRows.Add(new DataRow(new Dictionary<string, string>() {
                { "Field01", "Value02-01" },
                { "Field02", "Value02-02" },
                { "Field03", "Value02-03" },
                { "Field04", "Value02-04" },
            },
            new ObjectSecurity(new List<Core.Framework.Provisioning.Model.RoleAssignment>() {
                new Core.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal01",
                    RoleDefinition ="Read"
                },
                new Core.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal02",
                    RoleDefinition ="Contribute"
                }
                ,
                new Core.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal03",
                    RoleDefinition ="FullControl"
                }
            })
            {
                ClearSubscopes = false,
                CopyRoleAssignments = false
            }));
            list.DataRows.Add(new DataRow(new Dictionary<string, string>() {
                { "Field01", "Value03-01" },
                { "Field02", "Value03-02" },
                { "Field03", "Value03-03" },
                { "Field04", "Value03-04" },
            }));
#endregion

            var ca = new CustomAction()
            {
                Name = "SampleCustomAction",
                Description = "Just a sample custom action",
                Enabled = true,
                Group = "Samples",
                ImageUrl = "OneImage.png",
                Location = "Any",
                RegistrationId = "0x0101",
                RegistrationType = UserCustomActionRegistrationType.ContentType,
                Sequence = 100,
                ScriptBlock = "scriptblock",
                ScriptSrc = "script.js",
                Url = "http://somewhere.com/",
                Rights = new BasePermissions(),
                Title = "Sample Action",
                Remove = true,
                CommandUIExtension = XElement.Parse("<CommandUIExtension><customElement><!--Whateveryoulikehere--></customElement></CommandUIExtension>")
            };
            ca.Rights.Set(PermissionKind.AddListItems);
            list.UserCustomActions.Add(ca);

#region views
            list.Views.Add(new Core.Framework.Provisioning.Model.View()
            {
                SchemaXml = @"<View DisplayName=""View One"">
                  <ViewFields>
                    <FieldRef Name=""ID"" />
                    <FieldRef Name=""Title"" />
                    <FieldRef Name=""ProjectID"" />
                    <FieldRef Name=""ProjectName"" />
                    <FieldRef Name=""ProjectManager"" />
                    <FieldRef Name=""DocumentDescription"" />
                  </ViewFields>
                  <Query>
                    <Where>
                      <Eq>
                        <FieldRef Name=""ProjectManager"" />
                        <Value Type=""Text"">[Me]</Value>
                      </Eq>
                    </Where>
                  </Query>
                </View>"
            });
            list.Views.Add(new Core.Framework.Provisioning.Model.View()
            { 
                SchemaXml = @"<View DisplayName=""View Two"">
                  <ViewFields>
                    <FieldRef Name=""ID"" />
                    <FieldRef Name=""Title"" />
                    <FieldRef Name=""ProjectID"" />
                    <FieldRef Name=""ProjectName"" />
                  </ViewFields>
                </View>"
            });
#endregion

#region fieldrefs
            list.FieldRefs.Add(new FieldRef("ProjectID")
            {
                Id = new Guid("{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}"),
                DisplayName = "Project ID",
                Hidden = false,
                Required = true
            });
            list.FieldRefs.Add(new FieldRef("ProjectName")
            {
                Id = new Guid("{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}"),
                DisplayName = "Project Name",
                Hidden = true,
                Required = false
            });
            list.FieldRefs.Add(new FieldRef("ProjectManager")
            {
                Id = new Guid("{A5DE9600-B7A6-42DD-A05E-10D4F1500208}"),
                DisplayName = "Project Manager",
                Hidden = false,
                Required = true
            });
#endregion

#region folders
            var folder01 = new Core.Framework.Provisioning.Model.Folder("Folder01");
            var folder02 = new Core.Framework.Provisioning.Model.Folder("Folder02");
            folder01.Folders.Add(new Core.Framework.Provisioning.Model.Folder("Folder01.01",
                security: new ObjectSecurity(new List<Core.Framework.Provisioning.Model.RoleAssignment>() {
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal01",
                        RoleDefinition ="Read"
                    },
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal02",
                        RoleDefinition ="Contribute"
                    },
                    new Core.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal03",
                        RoleDefinition ="FullControl"
                    }
                })
                {
                    CopyRoleAssignments = true,
                    ClearSubscopes = true
                }));
            folder01.Folders.Add(new Core.Framework.Provisioning.Model.Folder("Folder01.02"));
            list.Folders.Add(folder01);
            list.Folders.Add(folder02);
#endregion

            list.Fields.Add(new Core.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"Project ID\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" Required=\"TRUE\" />"
            });
            list.Fields.Add(new Core.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID=\"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"Project Name\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });

            result.Lists.Add(list);

            var serializer = new XMLPnPSchemaV201605Serializer();
            provider.SaveAs(result, "ProvisioningTemplate-2016-05-Sample-03-OUT-lst.xml", serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningTemplate-2016-05-Sample-03-OUT-lst.xml";
            Assert.IsTrue(System.IO.File.Exists(path));
            XDocument xml = XDocument.Load(path);
            Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning wrappedResult =
                XMLSerializer.Deserialize<Core.Framework.Provisioning.Providers.Xml.V201605.Provisioning>(xml);

            Assert.IsNotNull(wrappedResult);
            Assert.IsNotNull(wrappedResult.Templates);
            Assert.AreEqual(1, wrappedResult.Templates.Count());
            Assert.IsNotNull(wrappedResult.Templates[0].ProvisioningTemplate);
            Assert.AreEqual(1, wrappedResult.Templates[0].ProvisioningTemplate.Count());

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.IsNotNull(template.Lists);
            Assert.AreEqual(1, template.Lists.Length);

            var l = template.Lists.FirstOrDefault(ls => ls.Title == "Project Documents");
            Assert.IsNotNull(l);
            Assert.IsTrue(l.ContentTypesEnabled);
            Assert.AreEqual("Project Documents are stored here", l.Description);
            Assert.AreEqual("document.dotx", l.DocumentTemplate);
            Assert.AreEqual(1, l.DraftVersionVisibility);
            Assert.IsTrue(l.DraftVersionVisibilitySpecified);
            Assert.IsTrue(l.EnableAttachments);
            Assert.IsTrue(l.EnableFolderCreation);
            Assert.IsTrue(l.EnableMinorVersions);
            Assert.IsTrue(l.EnableModeration);
            Assert.IsTrue(l.EnableVersioning);
            Assert.IsTrue(l.ForceCheckout);
            Assert.IsTrue(l.Hidden);
            Assert.AreEqual(10, l.MaxVersionLimit);
            Assert.IsTrue(l.MaxVersionLimitSpecified);
            Assert.AreEqual(2, l.MinorVersionLimit);
            Assert.IsTrue(l.MinorVersionLimitSpecified);
            Assert.IsTrue(l.OnQuickLaunch);
            Assert.IsTrue(l.RemoveExistingContentTypes);
            Assert.AreEqual("30FB193E-016E-45A6-B6FD-C6C2B31AA150".ToLower(), l.TemplateFeatureID);
            Assert.AreEqual(101, l.TemplateType);
            Assert.AreEqual("/Lists/ProjectDocuments", l.Url);

            Assert.IsNotNull(l.Security);
            var security = l.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            var ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            Assert.IsNotNull(list.ContentTypeBindings);
            Assert.AreEqual(3, list.ContentTypeBindings.Count);
            var ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeID == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.IsNotNull(ct);
            Assert.IsTrue(ct.Default);
            Assert.IsFalse(ct.Remove);
            ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeID == "0x0101");
            Assert.IsNotNull(ct);
            Assert.IsFalse(ct.Default);
            Assert.IsTrue(ct.Remove);
            ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeID == "0x0102");
            Assert.IsNotNull(ct);
            Assert.IsFalse(ct.Default);
            Assert.IsFalse(ct.Remove);

            Assert.IsNotNull(l.FieldDefaults);
            Assert.AreEqual(4, l.FieldDefaults.Length);
            var fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field01");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue01", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field02");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue02", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field03");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue03", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field04");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue04", fd.Value);

            Assert.IsNotNull(l.DataRows);
            Assert.AreEqual(3, l.DataRows.Length);
#region data row 1 asserts
            var dr = l.DataRows.FirstOrDefault(r => r.DataValue.Any(d => d.Value.StartsWith("Value01")));
            Assert.IsNotNull(dr);
            Assert.IsNotNull(dr.Security);
            security = dr.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            Assert.IsNotNull(dr.DataValue);
            var dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-01", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-02", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-03", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-04", dv.Value);
#endregion
#region data row 2 asserts
            dr = l.DataRows.FirstOrDefault(r => r.DataValue.Any(d => d.Value.StartsWith("Value02")));
            Assert.IsNotNull(dr);
            Assert.IsNotNull(dr.Security);
            security = dr.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsFalse(security.ClearSubscopes);
            Assert.IsFalse(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            Assert.IsNotNull(dr.DataValue);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-01", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-02", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-03", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-04", dv.Value);
#endregion
#region data row 3 asserts
            dr = l.DataRows.FirstOrDefault(r => r.DataValue.Any(d => d.Value.StartsWith("Value03")));
            Assert.IsNotNull(dr);
            Assert.IsNull(dr.Security);
            
            Assert.IsNotNull(dr.DataValue);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-01", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-02", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-03", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-04", dv.Value);
#endregion

#region user custom action
            Assert.IsNotNull(l.UserCustomActions);
            Assert.AreEqual(1, l.UserCustomActions.Length);
            var ua = l.UserCustomActions.FirstOrDefault(a => a.Name == "SampleCustomAction");
            Assert.IsNotNull(ua);
            Assert.AreEqual("Just a sample custom action", ua.Description);
            Assert.IsTrue(ua.Enabled);
            Assert.AreEqual("Samples", ua.Group);
            Assert.AreEqual("OneImage.png", ua.ImageUrl);
            Assert.AreEqual("Any", ua.Location);
            Assert.AreEqual("0x0101", ua.RegistrationId);
            Assert.AreEqual(Core.Framework.Provisioning.Providers.Xml.V201605.RegistrationType.ContentType, ua.RegistrationType);
            Assert.AreEqual(100, ua.Sequence);
            Assert.AreEqual("scriptblock", ua.ScriptBlock);
            Assert.AreEqual("script.js", ua.ScriptSrc);
            Assert.AreEqual("http://somewhere.com/", ua.Url);
            Assert.AreEqual("Sample Action", ua.Title);
            Assert.IsTrue(ua.Remove);
            Assert.IsNotNull(ua.CommandUIExtension);
            Assert.IsNotNull(ua.CommandUIExtension.Any);
            Assert.AreEqual(1, ua.CommandUIExtension.Any.Length);
            Assert.IsNotNull(ua.Rights);
            Assert.IsTrue(ua.Rights.Contains("AddListItems"));
#endregion

            Assert.IsNotNull(l.Views);
            Assert.IsNotNull(l.Views.Any);
            Assert.AreEqual(2, l.Views.Any.Length);

#region field refs
            Assert.IsNotNull(l.FieldRefs);
            Assert.AreEqual(3, l.FieldRefs.Length);
            var fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectID");
            Assert.IsNotNull(fr);
            Assert.AreEqual("23203E97-3BFE-40CB-AFB4-07AA2B86BF45".ToLower(), fr.ID);
            Assert.AreEqual("Project ID", fr.DisplayName);
            Assert.IsFalse(fr.Hidden);
            Assert.IsTrue(fr.Required);
            fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectName");
            Assert.IsNotNull(fr);
            Assert.AreEqual("B01B3DBC-4630-4ED1-B5BA-321BC7841E3D".ToLower(), fr.ID);
            Assert.AreEqual("Project Name", fr.DisplayName);
            Assert.IsTrue(fr.Hidden);
            Assert.IsFalse(fr.Required);
            fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectManager");
            Assert.IsNotNull(fr);
            Assert.AreEqual("A5DE9600-B7A6-42DD-A05E-10D4F1500208".ToLower(), fr.ID);
            Assert.AreEqual("Project Manager", fr.DisplayName);
            Assert.IsFalse(fr.Hidden);
            Assert.IsTrue(fr.Required);
#endregion

#region folders
            Assert.IsNotNull(l.Folders);
            Assert.AreEqual(2, l.Folders.Length);
            var fl = l.Folders.FirstOrDefault(f => f.Name == "Folder02");
            Assert.IsNotNull(fl);
            Assert.IsNull(fl.Folder1);
            fl = l.Folders.FirstOrDefault(f => f.Name == "Folder01");
            Assert.IsNotNull(fl);
            Assert.IsNotNull(fl.Folder1);
            var fl1 = fl.Folder1.FirstOrDefault(f=>f.Name == "Folder01.02");
            Assert.IsNotNull(fl1);
            Assert.IsNull(fl1.Folder1);
            fl1 = fl.Folder1.FirstOrDefault(f => f.Name == "Folder01.01");
            Assert.IsNull(fl1.Folder1);
            Assert.IsNotNull(fl1.Security);
            security = fl1.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);
#endregion

            Assert.IsNotNull(l.Fields);
            Assert.IsNotNull(l.Fields.Any);
            Assert.AreEqual(2, l.Fields.Any.Length);
            Assert.IsTrue(l.Fields.Any.All(x => x.OuterXml.StartsWith("<Field")));
        }
#endregion
    }
}
#endif