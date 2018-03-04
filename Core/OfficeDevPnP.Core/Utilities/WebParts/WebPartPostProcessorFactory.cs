using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using OfficeDevPnP.Core.Utilities.WebParts.Processors;

namespace OfficeDevPnP.Core.Utilities.WebParts
{
    /// <summary>
    /// Creates <see cref="IWebPartPostProcessor"/> by parsing web part schema xml
    /// </summary>
    public class WebPartPostProcessorFactory
    {
        /// <summary>
        /// Resolves webpart by parsing web part schema xml
        /// </summary>
        /// <param name="wpXml">WebPart schema xml</param>
        /// <returns>Returns PassThroughProcessor object</returns>
        public static IWebPartPostProcessor Resolve(string wpXml)
        {
            //don't care about web parts with old schema version (v2)
            if (wpXml.IndexOf("xmlns=\"http://schemas.microsoft.com/WebPart/v3\"", StringComparison.OrdinalIgnoreCase) == -1)
            {
                return new PassThroughProcessor();
            }

            var serializer = new XmlSerializer(typeof(Schema.WebParts));

            using (var xmlStream = GetXmlStream(wpXml))
            using (var xmlReader = new XmlTextReader(xmlStream))
            {
                xmlReader.Namespaces = false;
                var wpSchema = (Schema.WebParts)serializer.Deserialize(xmlReader);

                if (wpSchema.WebPart.MetaData.Type.Name.IndexOf("Microsoft.SharePoint.WebPartPages.XsltListViewWebPart", StringComparison.Ordinal) != -1)
                {
                    return new XsltWebPartPostProcessor(wpSchema.WebPart);
                }
            }

            return new PassThroughProcessor();
        }

        private static Stream GetXmlStream(string xml)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(xml);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
    }
}
