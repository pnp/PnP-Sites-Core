using OfficeDevPnP.Core.Framework.Provisioning.Connectors;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    public class XMLZipPackageTemplateProvider : XMLTemplateProvider
    {

        public XMLZipPackageTemplateProvider()
        {

        }
        public XMLZipPackageTemplateProvider(string connectionString, string container) :
            base(new ZipFileConnector(connectionString, container))
        {
        }
    }
}
