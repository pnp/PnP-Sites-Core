using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    public class XMLOpenXMLTemplateProvider : XMLTemplateProvider
    {
        public XMLOpenXMLTemplateProvider(string packageFileName,
            FileConnectorBase persistenceConnector,
            String author = null,
            X509Certificate2 signingCertificate = null) :
            base(new OpenXMLConnector(packageFileName, persistenceConnector,
                author, signingCertificate))
        {
        }

        public XMLOpenXMLTemplateProvider(OpenXMLConnector openXMLConnector) :
            base(openXMLConnector)
        {
        }
    }
}
