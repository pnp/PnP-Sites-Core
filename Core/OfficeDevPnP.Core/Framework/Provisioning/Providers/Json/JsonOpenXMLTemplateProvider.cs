using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Json
{
    public class JsonOpenXMLTemplateProvider : JsonTemplateProvider
    {
        public JsonOpenXMLTemplateProvider(string packageFileName,
            FileConnectorBase persistenceConnector,
            String author = null,
            X509Certificate2 signingCertificate = null) :
            base(new OpenXMLConnector(packageFileName, persistenceConnector,
                author, signingCertificate))
        {
        }

        public JsonOpenXMLTemplateProvider(OpenXMLConnector openXMLConnector) :
            base(openXMLConnector)
        {
        }
    }
}
