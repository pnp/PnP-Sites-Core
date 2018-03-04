#if !NETSTANDARD2_0
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors
{
    /// <summary>
    /// Connector for files in Azure blob storage
    /// </summary>
    public class AzureStorageConnector : FileConnectorBase
    {
        #region private variables
        private bool initialized = false;
        private CloudBlobClient blobClient = null;
        private const string FOLDER = "Folder";
        #endregion

        #region Constructor
        /// <summary>
        /// Base constructor
        /// </summary>
        public AzureStorageConnector() : base()
        {

        }

        /// <summary>
        /// AzureStorageConnector constructor. Allows to directly set Azure Storage key and container
        /// </summary>
        /// <param name="connectionString">Azure Storage Key (DefaultEndpointsProtocol=https;AccountName=yyyy;AccountKey=xxxx)</param>
        /// <param name="container">Name of the Azure container to operate against</param>
        public AzureStorageConnector(string connectionString, string container) : base()
        {
            if (String.IsNullOrEmpty(connectionString))
            {
                throw new ArgumentException("connectionString");
            }

            if (String.IsNullOrEmpty(container))
            {
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');
            
            this.AddParameterAsString(CONNECTIONSTRING, connectionString);
            this.AddParameterAsString(CONTAINER, container);
        }
        #endregion

        #region Base class overrides
        /// <summary>
        /// Get the files available in the default container
        /// </summary>
        /// <returns>List of files</returns>
        public override List<string> GetFiles()
        {
            return GetFiles(GetContainer());
        }

        /// <summary>
        /// Get the files available in the specified container
        /// </summary>
        /// <param name="container">Name of the container to get the files from</param>
        /// <returns>List of files</returns>
        public override List<string> GetFiles(string container)
        {
            if (String.IsNullOrEmpty(container))
            {
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');
            
            if (!initialized)
            {
                Initialize();
            }

            List<string> result = new List<string>();

            var containerTuple = ParseContainer(container);

            container = containerTuple.Item1;
            string prefix = string.IsNullOrEmpty(containerTuple.Item2) ? null : containerTuple.Item2;

            CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);

            foreach (IListBlobItem item in blobContainer.ListBlobs(prefix, false))
            {
                if (item.GetType() == typeof(CloudBlockBlob))
                {
                    CloudBlockBlob blob = (CloudBlockBlob)item;
                    result.Add(blob.Name);
                }
            }

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
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');
            
            if (!initialized)
            {
                Initialize();
            }

            List<string> result = new List<string>();

            var containerTuple = ParseContainer(container);

            container = containerTuple.Item1;
            string prefix = string.IsNullOrEmpty(containerTuple.Item2) ? null : containerTuple.Item2;

            CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);

            foreach (IListBlobItem item in blobContainer.ListBlobs(prefix, false))
            {
                if (item.GetType() == typeof(CloudBlobDirectory))
                {
                    CloudBlobDirectory blob = (CloudBlobDirectory)item;
                    result.Add(blob.Uri.ToString());
                }
            }

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
        /// Gets a file as string from the specified container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <param name="container">Name of the container to get the file from</param>
        /// <returns>String containing the file contents</returns>
        public override String GetFile(string fileName, string container)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (String.IsNullOrEmpty(container))
            {
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');            

            string result = null;
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
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');            

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
                throw new ArgumentException(nameof(container));
            }
            container = container.Replace('\\', '/');            

            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (!initialized)
            {
                Initialize();
            }

            try
            {
                var containerTuple = ParseContainer(container);

                container = containerTuple.Item1;
                fileName = string.Concat(containerTuple.Item2, fileName);

                CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);

                // Create the container if it doesn't already exist.
                blobContainer.CreateIfNotExists();

                CloudBlockBlob blockBlob = blobContainer.GetBlockBlobReference(fileName);

                blockBlob.UploadFromStream(stream);
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FileSaved, fileName, container);
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FileSaveFailed, fileName, container, ex.Message);
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
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');            

            if (!initialized)
            {
                Initialize();
            }

            try
            {
                CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);
                CloudBlockBlob blockBlob = blobContainer.GetBlockBlobReference(fileName);

                if (blockBlob.Exists())
                {
                    blockBlob.Delete();
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FileDeleted, fileName, container);
                }
                else
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FileDeleteNotFound, fileName, container);
                }
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FileDeleteFailed, fileName, container, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns a filename without a path
        /// </summary>
        /// <param name="fileName">Name of the file</param>
        /// <returns>Returns filename without path</returns>
        public override string GetFilenamePart(string fileName)
        {
            return Path.GetFileName(fileName);
        }

        #endregion

        #region Private methods
        private void Initialize()
        {
            try
            {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(GetConnectionString());
                blobClient = storageAccount.CreateCloudBlobClient();
                initialized = true;
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FailedToInitialize, ex.Message);
                throw;
            }
        }

        private MemoryStream GetFileFromStorage(string fileName, string container)
        {
            if (!initialized)
            {
                Initialize();
            }

            try
            {
                var containerTuple = ParseContainer(container);

                container = containerTuple.Item1;
                fileName = string.Concat(containerTuple.Item2, fileName);

                CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);
                CloudBlockBlob blockBlob = blobContainer.GetBlockBlobReference(fileName);

                MemoryStream result = new MemoryStream();
                blockBlob.DownloadToStream(result);
                result.Position = 0;

                Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FileRetrieved, fileName, container);
                return result;
            }
            catch (StorageException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_Azure_FileNotFound, fileName, container, ex.Message);
                return null;
            }
        }

        private Tuple<string, string> ParseContainer(string container)
        {
            var firstOccouranceOfSlash = container.IndexOf('/');
            var folder = string.Empty;

            if (firstOccouranceOfSlash > -1)
            {
                var orgContainer = container;
                container = orgContainer.Substring(0, firstOccouranceOfSlash);
                folder = orgContainer.Substring(firstOccouranceOfSlash + 1);
                if (!folder.Substring(folder.Length - 1, 1).Equals("/", StringComparison.InvariantCultureIgnoreCase))
                {
                    folder = folder + "/";
                }
            }

            return Tuple.Create(container, folder);
        }
        #endregion


    }
}
#endif