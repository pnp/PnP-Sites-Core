using OfficeDevPnP.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors
{
	/// <summary>
	/// Connector for zip files in file system
	/// </summary>
	public class ZipFileConnector : FileConnectorBase
	{
		#region Constructors
		/// <summary>
		/// Base constructor
		/// </summary>
		public ZipFileConnector()
			: base()
		{

		}

		/// <summary>
		/// ZipFileSystemConnector constructor. Allows to directly set zip file path and root folder
		/// </summary>
		/// <param name="connectionString">Path to zip file (e.g. c:\temp\sitetemplate.zip .\sitetemplate.zip)</param>
		/// <param name="container">Root folder inside zip file (e.g. templates or resources\templates or blank</param>
		public ZipFileConnector(string connectionString, string container)
			: base()
		{
			if (String.IsNullOrEmpty(connectionString))
			{
				throw new ArgumentException("connectionString");
			}

			if (String.IsNullOrEmpty(container))
			{
				container = "";
			}

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
			List<string> result = new List<string>();
			using (ZipArchive archive = OpenArchive())
			{
				foreach (ZipArchiveEntry entry in archive.Entries)
				{
					if (IsFileInContainer(container, entry.FullName))
					{
						result.Add(Path.GetFileName(entry.FullName));
					}
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

			string result = null;
			MemoryStream stream = null;
			try
			{
				stream = GetFileFromArchive(fileName, container);
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

			return GetFileFromArchive(fileName, container);
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
				throw new ArgumentException("fileName");
			}

			if (String.IsNullOrEmpty(container))
			{
				container = "";
			}

			if (stream == null)
			{
				throw new ArgumentNullException("stream");
			}

			try
			{
				container = FormatContainerName(container);
				string filePath = Path.Combine(container, fileName);

				using (var archive = OpenArchive(ZipArchiveMode.Update, true))
				{
					var entry = archive.Mode != ZipArchiveMode.Create ?
						archive.Entries.FirstOrDefault(e => string.Equals(e.FullName, filePath, StringComparison.InvariantCultureIgnoreCase)) :
						null;
					if (entry == null)
					{
						entry = archive.CreateEntry(filePath);
					}
					using (var entryStream = entry.Open())
					{
						byte[] buffer = new byte[16 * 1024];
						int read = 0;
						if (entryStream.CanSeek)
						{
							entryStream.SetLength(0);
						}
						stream.Seek(0, SeekOrigin.Begin);
						while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
						{
							entryStream.Write(buffer, 0, read);
						}
					}
				}
				Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileSaved, fileName, container);
			}
			catch (Exception ex)
			{
				Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileSaveFailed, fileName, container, ex.Message);
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
				container = FormatContainerName(container);
				string filePath = Path.Combine(container, fileName);

				using (var archive = OpenArchive(ZipArchiveMode.Update))
				{
					var entry = archive.Entries.FirstOrDefault(e => string.Equals(e.FullName, filePath, StringComparison.InvariantCultureIgnoreCase));
					if (entry != null)
					{
						entry.Delete();
						Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileDeleted, fileName, container);
					}
					else
					{
						Log.Warning(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileDeleteNotFound, fileName, container);
					}
				}
			}
			catch (Exception ex)
			{
				Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileDeleteFailed, fileName, container, ex.Message);
				throw;
			}
		}
		#endregion

		#region Private methods

		private ZipArchive OpenArchive(ZipArchiveMode mode = ZipArchiveMode.Read, bool createIfNotExist = false)
		{
			Stream stream = null;
			try
			{
				stream = new FileStream(GetConnectionString(), FileMode.Open);
			}
			catch (FileNotFoundException)
			{
				if (createIfNotExist)
				{
					mode = ZipArchiveMode.Create;
					stream = new FileStream(GetConnectionString(), FileMode.CreateNew);
				}
				else
				{
					throw;
				}
			}

			return new ZipArchive(stream, mode);
		}

		private string FormatContainerName(string container)
		{
			string res = container;
			if (!String.IsNullOrEmpty(container))
			{
				res = container.Replace("\\", "/");
				if (!container.EndsWith("/"))
				{
					res += "/";
				}
			}
			return res;
		}

		private bool IsFileInContainer(string container, string filePath)
		{
			bool res = false;
			var fileDir = Path.GetDirectoryName(filePath);
            if (String.IsNullOrEmpty(container) && String.IsNullOrEmpty(fileDir))
			{
				res = true;
			}
			else if (!String.IsNullOrEmpty(filePath))
			{
				container = FormatContainerName(container);
				fileDir = FormatContainerName(fileDir);

				res = string.Equals(container, fileDir, StringComparison.CurrentCultureIgnoreCase);
			}
			return res;
		}

		private MemoryStream GetFileFromArchive(string fileName, string container)
		{
			container = FormatContainerName(container);
			string filePath = Path.Combine(container, fileName);

			MemoryStream stream = null;
			try
			{
				using (var archive = OpenArchive())
				{
					var entry = archive.Entries.FirstOrDefault(e => string.Equals(e.FullName, filePath, StringComparison.InvariantCultureIgnoreCase));
					if (entry != null)
					{
						byte[] buffer = new byte[16 * 1024];
						using (var entryStream = entry.Open())
						{
							stream = new MemoryStream();
							int read = 0;
							while ((read = entryStream.Read(buffer, 0, buffer.Length)) > 0)
							{
								stream.Write(buffer, 0, read);
							}
							Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileRetrieved, fileName, container);
							stream.Position = 0;
						}
					}
				}
			}
			catch
			{
				if (stream != null)
				{
					stream.Dispose();
				}
				throw;
			}
			return stream;
		}
		#endregion
	}
}
