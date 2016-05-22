using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System.IO;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

namespace OfficeDevPnP.Core.Tests.Framework.Connectors
{
    [TestClass]
    public class ConnectorOpenXmlTests
    {
        private const string packageFileName = "TestTemplate.pnp";

        #region Test initialize and cleanup

        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            // File system setup
            if (System.IO.File.Exists(@".\Resources\Templates\TestTemplate.pnp"))
            {
                System.IO.File.Delete(@".\Resources\Templates\TestTemplate.pnp");
            }
        }

        #endregion

        #region OpenXML Connector tests

        /// <summary>
        /// Create a PNP OpenXML package file and add a sample template to it
        /// </summary>
        [TestMethod]
        public void OpenXMLSaveTemplate()
        {
            OpenXMLSaveTemplateInternal();
            Boolean checkFileExistence = File.Exists(String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory)
                    + @"\Templates\TestTemplate.pnp");
            Assert.IsTrue(checkFileExistence);
        }

        private void OpenXMLSaveTemplateInternal()
        {
            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector,
                "OfficeDevPnP Automated Test");

            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\Templates\ProvisioningSchema-2015-12-FullSample-02.xml", "", openXMLConnector);
            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagelogo.png", "Images", openXMLConnector);
            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagebg.jpg", "Images", openXMLConnector);

            if (openXMLConnector is ICommitableFileConnector)
            {
                ((ICommitableFileConnector)openXMLConnector).Commit();
            }
        }

        [TestMethod]
        public void OpenXMLLoadTemplate()
        {
            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector);

            var templateFile = openXMLConnector.GetFileStream("ProvisioningSchema-2015-12-FullSample-02.xml");

            XMLPnPSchemaV201512Formatter formatter = new XMLPnPSchemaV201512Formatter();
            Boolean checkTemplate = formatter.IsValid(templateFile);

            Assert.IsTrue(checkTemplate);

            var image1 = openXMLConnector.GetFileStream("garagelogo.png", "Images");
            Assert.IsNotNull(image1);

            var image2 = openXMLConnector.GetFileStream("garagelogo.png", "Images");
            Assert.IsNotNull(image2);
        }

        [TestMethod]
        public void OpenXMLDeleteFileFromTemplate()
        {
            OpenXMLSaveTemplateInternal();

            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector);

            openXMLConnector.DeleteFile("garagelogo.png", "Images");

            var image1 = openXMLConnector.GetFileStream("garagelogo.png", "Images");
            Assert.IsNull(image1);
        }

        private static void SaveFileInPackage(String filePath, String container, FileConnectorBase connector)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                String fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
                connector.SaveFileStream(fileName, container, fs);
            }
        }

        #endregion
    }
}
