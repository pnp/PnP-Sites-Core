using OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML
{
    /// <summary>
    /// Extension class for PnP OpenXML package files
    /// </summary>
    public static class PnPPackageExtensions
    {

        public static MemoryStream PackTemplateAsStream(this PnPInfo pnpInfo)
        {
            MemoryStream stream = new MemoryStream();
            using (PnPPackage package = PnPPackage.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                SavePnPPackage(pnpInfo, package);
            }
            stream.Position = 0;
            return stream;
        }

        public static Byte[] PackTemplate(this PnPInfo pnpInfo)
        {
            using (MemoryStream stream = PackTemplateAsStream(pnpInfo))
            {
                return stream.ToArray();
            }
        }

        public static PnPInfo UnpackTemplate(this MemoryStream stream)
        {
            PnPInfo siteTemplate;
            using (PnPPackage package = PnPPackage.Open(stream, FileMode.Open, 
                stream.CanWrite ? FileAccess.ReadWrite : FileAccess.Read))
            {
                siteTemplate = LoadPnPPackage(package);
            }
            return siteTemplate;
        }

        public static PnPInfo UnpackTemplate(this Byte[] packageBytes)
        {
            using (MemoryStream stream = new MemoryStream(packageBytes))
            {
                return UnpackTemplate(stream);
            }
        }

        public static string AsInternalFilename(this string filename)
        {
            return Guid.NewGuid() + Path.GetExtension(filename);
        }

        #region Private Methods for handling templates

        private static PnPInfo LoadPnPPackage(PnPPackage package)
        {
            PnPInfo pnpInfo = new PnPInfo();
            pnpInfo.Manifest = package.Manifest;
            pnpInfo.Properties = package.Properties;
            pnpInfo.FilesMap = package.FilesMap;

            pnpInfo.Files = new List<PnPFileInfo>();

            foreach (KeyValuePair<String, PnPPackageFileItem> file in package.Files)
            {
                pnpInfo.Files.Add(
                    new PnPFileInfo
                    {
                        InternalName = file.Key,
                        OriginalName = package.FilesMap != null ?
                            (String.IsNullOrEmpty(file.Value.Folder) ?
                            package.FilesMap.Map[file.Key] :
                            package.FilesMap.Map[file.Key].Replace(file.Value.Folder + '/', "")) :
                            file.Key,
                        Folder = file.Value.Folder,
                        Content = file.Value.Content,
                    });
            }
            return pnpInfo;
        }

        private static void SavePnPPackage(PnPInfo pnpInfo, PnPPackage package)
        {
            package.Manifest = pnpInfo.Manifest;
            package.Properties = pnpInfo.Properties;
            Debug.Assert(pnpInfo.Files.TrueForAll(f => !string.IsNullOrWhiteSpace(f.InternalName)), "All files need an InternalFileName");
            package.FilesMap = new PnPFilesMap(pnpInfo.Files.ToDictionary(f => f.InternalName, f => Path.Combine(f.Folder, f.OriginalName).Replace('\\', '/').TrimStart('/')));
            package.ClearFiles();
            if (pnpInfo.Files != null)
            {
                foreach (PnPFileInfo file in pnpInfo.Files)
                {
                    package.AddFile(file.InternalName, file.Content);
                }
            }
        }

        #endregion
    }
}
