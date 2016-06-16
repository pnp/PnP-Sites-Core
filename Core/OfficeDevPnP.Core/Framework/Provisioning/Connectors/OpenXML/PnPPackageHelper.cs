using OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model;
using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML
{
    /// <summary>
    /// Extension class for PnP OpenXML package files
    /// </summary>
    public static class PnPPackageExtensions
    {
        public static MemoryStream PackTemplate(this PnPInfo pnpInfo)
        {
            MemoryStream stream = new MemoryStream();
            using (PnPPackage package = PnPPackage.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                SavePnPPackage(pnpInfo, package);
            }
            stream.Position = 0;
            return stream;

        }

        public static PnPInfo UnpackTemplate(this MemoryStream stream)
        {
            PnPInfo siteTemplate;
            using (PnPPackage package = PnPPackage.Open(stream, FileMode.Open, FileAccess.Read))
            {
                siteTemplate = LoadPnPPackage(package);
            }
            return siteTemplate;
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
                        OriginalName = package.FilesMap.Map[file.Key],
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
            package.FilesMap = new PnPFilesMap(pnpInfo.Files.ToDictionary(f => f.InternalName, f => f.OriginalName));
            package.ClearFiles();
            if (pnpInfo.Files != null)
            {
                foreach (PnPFileInfo file in pnpInfo.Files)
                {
                    package.AddFile(file.InternalName, file.Folder, file.Content);
                }
            }
        }

        #endregion
    }
}
