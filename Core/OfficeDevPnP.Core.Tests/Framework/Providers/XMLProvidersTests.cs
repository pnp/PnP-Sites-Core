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

        #region Formatter Refactoring Tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_201605()
        {
            var _expectedID = "SPECIALTEAM-01";
            var _expectedVersion = 1.0;

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
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents.Trim());

            Assert.IsNotNull(file.WebParts);
            webpart = file.WebParts.FirstOrDefault(wp => wp.Title == "My Editor");
            Assert.IsNotNull(webpart);
            Assert.AreEqual((uint)10, webpart.Order);
            Assert.AreEqual("Left", webpart.Zone);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents.Trim());

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
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents);

            Assert.IsNotNull(page.WebParts);
            webpart = page.WebParts.FirstOrDefault(wp => wp.Title == "My Editor");
            Assert.IsNotNull(webpart);
            Assert.AreEqual((uint)2, webpart.Row);
            Assert.AreEqual((uint)1, webpart.Column);
            Assert.IsNotNull(webpart.Contents);
            Assert.AreEqual("<webPart>[!<![CDATA[web part definition goes here]]></webPart>", webpart.Contents);

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
            Assert.AreEqual(0, tm.Language);
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
            Assert.IsNull(group.TermSets);
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
            var handler = template.ExtensibilityHandlers.FirstOrDefault(p => p.Type == "MyType1");
            Assert.IsNotNull(handler);
            Assert.AreEqual("MyAssembly1", handler.Assembly);
            Assert.IsTrue(handler.Enabled);
            Assert.AreEqual("<TestConfiguration xmlns=\"MyHandler\">Value</TestConfiguration>", handler.Configuration.Trim());

            handler = template.ExtensibilityHandlers.FirstOrDefault(p => p.Type == "MyType2");
            Assert.IsNotNull(handler);
            Assert.AreEqual("MyAssembly2", handler.Assembly);
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
            var handler = template.Providers.FirstOrDefault(p => p.HandlerType == "MyType1, MyAssembly1");
            Assert.IsNotNull(handler);
            Assert.IsTrue(handler.Enabled);
            Assert.AreEqual("<TestConfiguration>Value</TestConfiguration>", handler.Configuration);

            handler = template.Providers.FirstOrDefault(p => p.HandlerType == "MyType2, MyAssembly2");
            Assert.IsNotNull(handler);
            Assert.IsFalse(handler.Enabled);
            Assert.IsNull(handler.Configuration.OuterXml);
        }
        #endregion
    }
}
