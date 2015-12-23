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
        public static String GetPublishingXmlTemplate(bool isOffice365, bool isProviderHosted, PublishingTypes publishingType)
        {
            string publishingXmlTemplate = string.Format("pubxml.{0}.{1}.{2}.xml", isOffice365 ? "office365" : "onpremises", isProviderHosted ? "providerhosted" : "sharepointhosted", publishingType.ToString()).ToLower();

            using (Stream stream = typeof(ResourceManager).Assembly.GetManifestResourceStream(String.Format("{0}.{1}", typeof(ResourceManager).Namespace, publishingXmlTemplate)))
            {
                StreamReader reader = new StreamReader(stream);
                return reader.ReadToEnd();
            }
        }

        public static String GetAssemblyDirectory()
        {
            string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return Path.GetDirectoryName(path);
        }

    }
}
