using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Security.Cryptography.X509Certificates;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    public class XMLOpenXMLTemplateProvider : XMLTemplateProvider
    {
        public XMLOpenXMLTemplateProvider(string packageFileName,
            FileConnectorBase persistenceConnector,
            String author = null,
            X509Certificate2 signingCertificate = null,
            string templateFileName = null) :
            base(new OpenXMLConnector(packageFileName, persistenceConnector,
                author, signingCertificate, templateFileName))
        {
        }

        public XMLOpenXMLTemplateProvider(OpenXMLConnector openXMLConnector) :
            base(openXMLConnector)
        {
        }

        public ProvisioningHierarchy GetHierarchy()
        {
            ProvisioningHierarchy result = null;

            var openXmlConnection = this.Connector as OpenXMLConnector;
            var fileName = openXmlConnection.Info.Properties.TemplateFileName;
            if (!String.IsNullOrEmpty(fileName))
            {


                var stream = this.Connector.GetFileStream(fileName);

                if (stream != null)
                {
                    var formatter = new XMLPnPSchemaFormatter();

                    ITemplateFormatter specificFormatter = formatter.GetSpecificFormatterInternal(ref stream);
                    specificFormatter.Initialize(this);
                    result = ((IProvisioningHierarchyFormatter)specificFormatter).ToProvisioningHierarchy(stream);
                }
            }
            return (result);
        }

    }
}
