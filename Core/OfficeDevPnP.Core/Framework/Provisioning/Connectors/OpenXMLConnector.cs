using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors
{
    /// <summary>
    /// Connector that stores all the files into a unique .PNP OpenXML package
    /// </summary>
    public class OpenXMLConnector : FileConnectorBase, ICommitableFileConnector
    {
        private readonly PnPInfo pnpInfo;
        private readonly FileConnectorBase persistenceConnector;
        private readonly String packageFileName;

        #region Constructors

        /// <summary>
        /// OpenXMLConnector constructor. Allows to manage a .PNP OpenXML package through an in memory stream.
        /// </summary>
        /// <param name="packageStream">The package stream</param>
        public OpenXMLConnector(Stream packageStream): base()
        {
            if (packageStream == null)
            {
                throw new ArgumentNullException(nameof(packageStream));
            }

            if (!packageStream.CanRead)
            {
                throw new ArgumentException("package");
            }

            // If the .PNP package exists unpack it into PnP OpenXML package info object
            MemoryStream ms = packageStream.ToMemoryStream();
            this.pnpInfo = ms.UnpackTemplate();
        }

        /// <summary>
        /// OpenXMLConnector constructor. Allows to manage a .PNP OpenXML package file through a supporting persistence connector.
        /// </summary>
        /// <param name="packageFileName">The name of the .PNP package file. If the .PNP extension is missing, it will be added</param>
        /// <param name="persistenceConnector">The FileConnector object that will be used for physical persistence of the file</param>
        /// <param name="author">The Author of the .PNP package file, if any. Optional</param>
        /// <param name="signingCertificate">The X.509 certificate to use for digital signature of the template, optional</param>
        public OpenXMLConnector(string packageFileName,
            FileConnectorBase persistenceConnector,
            String author = null,
            X509Certificate2 signingCertificate = null)
            : base()
        {
            if (String.IsNullOrEmpty(packageFileName))
            {
                throw new ArgumentException("packageFileName");
            }
            if (!packageFileName.EndsWith(".pnp", StringComparison.InvariantCultureIgnoreCase))
            {
                // Check for file extension
                packageFileName = packageFileName.TrimEnd('.') + ".pnp";
            }

            this.packageFileName = packageFileName;

            if (persistenceConnector == null)
            {
                throw new ArgumentException("persistenceConnector");
            }
            this.persistenceConnector = persistenceConnector;

            // Try to load the .PNP package file
            var packageStream = persistenceConnector.GetFileStream(packageFileName);
            if (packageStream != null)
            {
                // If the .PNP package exists unpack it into PnP OpenXML package info object
                MemoryStream ms = packageStream.ToMemoryStream();
                this.pnpInfo = ms.UnpackTemplate();
            }
            else
            {
                // Otherwsie initialize a fresh new PnP OpenXML package info object
                this.pnpInfo = new PnPInfo()
                {
                    Manifest = new PnPManifest()
                    {
                        Type = PackageType.Full
                    },
                    Properties = new PnPProperties()
                    {
                        Generator = PnPCoreUtilities.PnPCoreVersionTag,
                        Author = !String.IsNullOrEmpty(author) ? author : String.Empty,
                    },
                };
            }
        }

        #endregion

        #region Base class overrides

        /// <summary>
        /// Get the files available in the default container
        /// </summary>
        /// <returns>List of files</returns>
        public override List<String> GetFiles()
        {
            return GetFiles(GetContainer());
        }

        /// <summary>
        /// Get the files available in the specified container
        /// </summary>
        /// <param name="container">Name of the container to get the files from (something like: "\images\subfolder")</param>
        /// <returns>List of files</returns>
        public override List<string> GetFiles(string container)
        {
            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }
			container = container.Replace(@"\", @"/").Trim('/');
			var result = (from file in this.pnpInfo.Files
                          where string.Equals(file.Folder, container, StringComparison.OrdinalIgnoreCase)
                          select file.OriginalName).ToList();

            return result;
        }

        /// <summary>
        /// Get the folders of the default container
        /// </summary>
        /// <returns>List of folders</returns>
        public override List<string> GetFolders()
        {
            return GetFolders(GetContainer());
        }

        /// <summary>
        /// Get the folders of a specified container
        /// </summary>
        /// <param name="container">Name of the container to get the folders from</param>
        /// <returns>List of folders</returns>
        public override List<string> GetFolders(string container)
        {
            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }

            var result = (from file in this.pnpInfo.Files
                          where file.Folder.StartsWith(container, StringComparison.InvariantCultureIgnoreCase) 
                            && !file.Folder.Equals(container, StringComparison.InvariantCultureIgnoreCase)
                          select file.Folder).ToList();

            return result;
        }

        /// <summary>
        /// Gets a file as string from the default container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <returns>String containing the file contents</returns>
        public override string GetFile(string fileName)
        {
            return GetFile(fileName, GetContainer());
        }

        /// <summary>
        /// Returns a filename without a path
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <returns>Returns a filename without a path</returns>
        public override string GetFilenamePart(string fileName)
        {
            return Path.GetFileName(fileName);
        }

        /// <summary>
        /// Gets a file as string from the specified container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <param name="container">Name of the container to get the file from</param>
        /// <returns>String containing the file contents</returns>
        public override string GetFile(string fileName, string container)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }

            String result = null;
            MemoryStream stream = null;
            try
            {
                stream = GetFileFromStorage(fileName, container);

                if (stream == null)
                {
                    return null;
                }

                result = Encoding.UTF8.GetString(stream.ToArray());
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }

            return result;
        }

        /// <summary>
        /// Gets a file as stream from the default container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <returns>String containing the file contents</returns>
        public override Stream GetFileStream(string fileName)
        {
            return GetFileStream(fileName, GetContainer());
        }

        /// <summary>
        /// Gets a file as stream from the specified container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <param name="container">Name of the container to get the file from</param>
        /// <returns>String containing the file contents</returns>
        public override Stream GetFileStream(string fileName, string container)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }

            return GetFileFromStorage(fileName, container);
        }

        /// <summary>
        /// Saves a stream to the default container with the given name. If the file exists it will be overwritten
        /// </summary>
        /// <param name="fileName">Name of the file to save</param>
        /// <param name="stream">Stream containing the file contents</param>
        public override void SaveFileStream(string fileName, Stream stream)
        {
            SaveFileStream(fileName, GetContainer(), stream);
        }

        /// <summary>
        /// Saves a stream to the specified container with the given name. If the file exists it will be overwritten
        /// </summary>
        /// <param name="fileName">Name of the file to save</param>
        /// <param name="container">Name of the container to save the file to</param>
        /// <param name="stream">Stream containing the file contents</param>
        public override void SaveFileStream(string fileName, string container, Stream stream)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException(nameof(fileName));
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }

			container = container.Replace(@"\", "/").Trim('/');

			if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            try
            {
                var memoryStream = stream.ToMemoryStream();
                byte[] bytes = memoryStream.ToArray();

                // Check if the file already exists in the package
                var existingFile = pnpInfo.Files.FirstOrDefault(f => f.OriginalName.Equals(fileName, StringComparison.InvariantCultureIgnoreCase) && f.Folder.Equals(container, StringComparison.InvariantCultureIgnoreCase));
                if (existingFile != null)
                {
                    existingFile.Content = bytes;
                }
                else
                {
                    pnpInfo.Files.Add(new PnPFileInfo
                    {
                        InternalName = fileName.AsInternalFilename(),
                        OriginalName = fileName,
                        Folder = container,
                        Content = bytes,
                    });
                }

                Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_OpenXML_FileSaved, fileName, container);
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_OpenXML_FileSaveFailed, fileName, container, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Deletes a file from the default container
        /// </summary>
        /// <param name="fileName">Name of the file to delete</param>
        public override void DeleteFile(string fileName)
        {
            DeleteFile(fileName, GetContainer());
        }

        /// <summary>
        /// Deletes a file from the specified container
        /// </summary>
        /// <param name="fileName">Name of the file to delete</param>
        /// <param name="container">Name of the container to delete the file from</param>
        public override void DeleteFile(string fileName, string container)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }

            try
            {
                var file = GetFileFromInsidePackage(fileName, container);
                if (file != null)
                {
                    pnpInfo.Files.Remove(file);
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_OpenXML_FileDeleted, fileName, container);
                }
                else
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_OpenXML_FileDeleteNotFound, fileName, container);
                }
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_OpenXML_FileDeleteFailed, fileName, container, ex.Message);
                throw;
            }
        }
        #endregion

        #region Private methods
        private MemoryStream GetFileFromStorage(string fileName, string container)
        {
            try
            {
                var file = GetFileFromInsidePackage(fileName, container);

                if (file != null)
                {
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_OpenXML_FileRetrieved, fileName, container);
                    var stream = new MemoryStream(file.Content);
                    return stream;
                }
                else
                {
                    throw new FileNotFoundException();
                }
            }
            catch (Exception ex)
            {
                if (ex is FileNotFoundException || ex is DirectoryNotFoundException)
                {
                    Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_OpenXML_FileNotFound, fileName, container, ex.Message);
                    return null;
                }

                throw;
            }
        }

        /// <summary>
        /// Will first try to find the file based on container/filename from the mapped file names.
        /// As a fallback it will try to find by container/filename in the pnp file structure, which was the original format.
        /// </summary>
        private PnPFileInfo GetFileFromInsidePackage(string fileName, string container)
        {
            string mappedPath = Path.Combine(container, fileName).Replace('\\', '/');
			PnPFileInfo file = null;
			if (pnpInfo.FilesMap != null)
			{
				 file = (from item in pnpInfo.FilesMap.Map
						 where item.Value.Equals(mappedPath, StringComparison.InvariantCultureIgnoreCase)
						 select pnpInfo.Files.FirstOrDefault(f => f.InternalName == item.Key)).FirstOrDefault();
			}
            if (file != null) return file;
            return pnpInfo.Files.FirstOrDefault(f => f.OriginalName.Equals(fileName, StringComparison.InvariantCultureIgnoreCase) && f.Folder.Equals(container, StringComparison.InvariantCultureIgnoreCase));
        }

        internal override string GetContainer()
        {
            // The is no default container
            return (String.Empty);
        }
        #endregion

        #region Commit capability

        /// <summary>
        /// Commits the file
        /// </summary>
        public void Commit()
        {
            MemoryStream stream = pnpInfo.PackTemplateAsStream();
            persistenceConnector.SaveFileStream(this.packageFileName, stream);
        }

        #endregion
    }
}
