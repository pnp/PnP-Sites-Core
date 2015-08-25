using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions.Resources
{
    public static class ResourceManager
    {
        public static Stream GetPublishingXmlTemplate(bool isOffice365, bool isProviderHosted, PublishingTypes publishingType)
        {
            string publishingXmlTemplate = string.Format("pubxml.{0}.{1}.{2}.xml", isOffice365 ? "office365" : "onpremises", isProviderHosted ? "providerhosted" : "sharepointhosted", publishingType.ToString()).ToLower();
            Stream stream = typeof(ResourceManager).Assembly.GetManifestResourceStream(publishingXmlTemplate);
            return stream;
        }

    }
}
