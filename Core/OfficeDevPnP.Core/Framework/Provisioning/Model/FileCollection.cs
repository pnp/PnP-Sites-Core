using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of File objects
    /// </summary>
    public partial class FileCollection : ProvisioningTemplateCollection<File>
    {
        public FileCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }

        public void AddFromResourcesFolder(string resourcesFolderPath, string destinationFolder, bool overwrite, bool recursive)
        {
            DirectoryInfo resourcesDirectoryInfo = new DirectoryInfo(resourcesFolderPath);

            foreach (var resourceFile in resourcesDirectoryInfo.GetFiles("*.*"))
            {
                File file = new File
                {
                    Folder = destinationFolder.Replace("\\", "/"),
                    Overwrite = overwrite,
                    Src = resourceFile.FullName
                };

                Add(file);
            }

            if (recursive)
            {
                foreach (var subDirectory in resourcesDirectoryInfo.GetDirectories())
                {
                    string subDirectoryDestinationFolder = Path.Combine(destinationFolder, subDirectory.Name);

                    AddFromResourcesFolder(subDirectory.FullName, subDirectoryDestinationFolder, true, true);
                }
            }
        }

        public void RemoveUnmodifiedSinceDate(DateTime date)
        {
            RemoveUnmodifiedSinceDate(this.ToList(), date);
        }

        public void RemoveUnmodifiedSinceDate(DateTime date, string folder)
        {
            // Only check for files to remove where target folder matches specified folder...
            RemoveUnmodifiedSinceDate(this.ToList().Where(f => f.Folder.Contains(folder)), date);
        }

        private void RemoveUnmodifiedSinceDate(IEnumerable<File> files, DateTime date)
        {
            int unmodifiedFilesCounter = 0;

            foreach (var file in files)
            {
                FileInfo fileInfo = new FileInfo(file.Src);

                if (fileInfo.LastWriteTimeUtc <= date.ToUniversalTime() &&
                    fileInfo.CreationTimeUtc <= date.ToUniversalTime())
                {
                    Remove(file);

                    unmodifiedFilesCounter++;
                }
            }

            Trace.TraceInformation($"{unmodifiedFilesCounter} unmodified files removed from Provisioning Template");
        }
    }
}
