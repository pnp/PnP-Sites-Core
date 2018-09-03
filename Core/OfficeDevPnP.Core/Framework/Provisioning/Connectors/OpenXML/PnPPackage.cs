using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Packaging;
using System.IO;
#if !NETSTANDARD2_0
using System.Xaml;
#else 
using Portable.Xaml;
#endif
using OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML
{
    /// <summary>
    /// Defines a PnP OpenXML package file
    /// </summary>
    public partial class PnPPackage : IDisposable
    {
#region Constant strings

        // site template xaml
        public const string R_PROVISIONINGTEMPLATE_MANIFEST = "http://schemas.dev.office.com/pnp/provisioningtemplate/v1/manifest";
        public const string R_PROVISIONINGTEMPLATE_BODY = "http://schemas.dev.office.com/pnp/provisioningtemplate/v1/body";
        public const string R_PROVISIONINGTEMPLATE_PROPERTIES = "http://schemas.dev.office.com/pnp/provisioningtemplate/v1/properties";
        public const string R_PROVISIONINGTEMPLATE_FILES_MAP = "http://schemas.dev.office.com/pnp/provisioningtemplate/v1/files.map";

        // supporting files
        public const string R_PROVISIONINGTEMPLATE_FILES_ORIGIN = "http://schemas.dev.office.com/pnp/provisioningtemplate/v1/files.origin";
        public const string R_PROVISIONINGTEMPLATE_FILE = "http://schemas.dev.office.com/pnp/provisioningtemplate/v1/file";

        // content types
        public const string CT_PROVISIONINGTEMPLATE_MANIFEST = "application/pnpprovisioningtemplate.manifest";
        public const string CT_PROVISIONINGTEMPLATE_BODY = "application/pnpprovisioningtemplate.body";
        public const string CT_PROVISIONINGTEMPLATE_PROPERTIES = "application/pnpprovisioningtemplate.properties";
        public const string CT_PROVISIONINGTEMPLATE_FILES_MAP = "application/pnpprovisioningtemplate.files.map";
        public const string CT_ORIGIN = "application/pnpprovisioningtemplate.origin";
        public const string CT_FILE = "application/unknown";

        // URLs
        public static string U_PROVISIONINGTEMPLATE_MANIFEST = "/manifest.xml";

        public static string U_DIR_PROVISIONINGTEMPLATE = "/ProvisioningTemplate/";
        public static string U_PROVISIONINGTEMPLATE_PROPERTIES = U_DIR_PROVISIONINGTEMPLATE + "props.xml";
        public static string U_PROVISIONINGTEMPLATE_FILES_MAP = U_DIR_PROVISIONINGTEMPLATE + "files-map.xml";
        public static string U_FILES_ORIGIN = "/files.origin";
        public static string U_DIR_FILES = "/Files/";

        // file extensions
        public const string EXT_PROVISIONINGTEMPLATE = ".pnp";

#endregion

        public const CompressionOption PACKAGE_COMPRESSION_LEVEL = CompressionOption.Maximum;

#region Public Properties

        /// <summary>
        /// The complete package object
        /// </summary>
        public Package Package { get; private set; }

        /// <summary>
        /// The Manifest Part of the package file
        /// </summary>
        public PackagePart ManifestPart
        {
            get
            {
                return GetSinglePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_MANIFEST, null);
            }
        }

        /// <summary>
        /// The Manifest of the package file
        /// </summary>
        public PnPManifest Manifest
        {
            get
            {
                return GetXamlSerializedPackagePartValue<PnPManifest>(ManifestPart);
            }
            set
            {
                PackagePart manifestPart = EnsurePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_MANIFEST, CT_PROVISIONINGTEMPLATE_MANIFEST, U_PROVISIONINGTEMPLATE_MANIFEST, null);
                SetXamlSerializedPackagePartValue(value, manifestPart);
            }
        }

        /// <summary>
        /// The Properties of the package
        /// </summary>
        public PnPProperties Properties
        {
            get
            {
                PackagePart propsPart = GetSinglePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_PROPERTIES, ManifestPart);
                return GetXamlSerializedPackagePartValue<PnPProperties>(propsPart);
            }
            set
            {
                PackagePart propsPart = EnsurePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_PROPERTIES, CT_PROVISIONINGTEMPLATE_PROPERTIES, U_PROVISIONINGTEMPLATE_PROPERTIES, ManifestPart);
                SetXamlSerializedPackagePartValue(value, propsPart);
            }
        }

        /// <summary>
        /// The File Map for files stored in the OpenXML file
        /// </summary>
        public PnPFilesMap FilesMap
        {
            get
            {
                PackagePart propsPart = GetSinglePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_FILES_MAP, ManifestPart);
                return GetXamlSerializedPackagePartValue<PnPFilesMap>(propsPart);
            }
            set
            {
                PackagePart propsPart = EnsurePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_FILES_MAP, CT_PROVISIONINGTEMPLATE_FILES_MAP, U_PROVISIONINGTEMPLATE_FILES_MAP, ManifestPart);
                SetXamlSerializedPackagePartValue(value, propsPart);
            }
        }

        /// <summary>
        /// The Files origin
        /// </summary>
        public PackagePart FilesOriginPart
        {
            get
            {
                return GetSinglePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_FILES_ORIGIN, ManifestPart);
            }
        }

        /// <summary>
        /// The Files Parts of the package
        /// </summary>
        public IList<PackagePart> FilesPackageParts
        {
            get
            {
                return GetAllPackagePartsWithRelationshipType(R_PROVISIONINGTEMPLATE_FILE, FilesOriginPart);
            }
        }

        /// <summary>
        /// The Files of the package
        /// </summary>
        public IDictionary<String, PnPPackageFileItem> Files 
        {
            get
            {
                Dictionary<String, PnPPackageFileItem> result = new Dictionary<String, PnPPackageFileItem>();
				var map = FilesMap?.Map;
                List<PackagePart> fileParts = GetAllPackagePartsWithRelationshipType(R_PROVISIONINGTEMPLATE_FILE, FilesOriginPart);
                foreach (PackagePart p in fileParts)
                {
					String fileName = p.Uri.ToString().Remove(0, U_DIR_FILES.Length);
					String folder = fileName.LastIndexOf('/') >= 0 ?
						fileName.Substring(0, fileName.LastIndexOf('/')) : String.Empty;

					fileName = fileName.Substring(fileName.LastIndexOf('/') + 1);

					if (string.IsNullOrEmpty(folder) && map != null && map.ContainsKey(fileName))
					{
						//try get folder from original name
						var originalName = map[fileName].Replace(@"\", "/");
						folder = originalName.LastIndexOf('/') >= 0 ?
							originalName.Substring(0, originalName.LastIndexOf('/')) : String.Empty;
					}

					Byte[] content = ReadPackagePartBytes(p);

                    result[fileName] = new PnPPackageFileItem
                    {
                        Name = fileName,
                        Folder = folder,
                        Content = content,
                    };
                }
                return result;
            }
        }

#endregion

#region Package Handling methods

        /// <summary>
        /// Opens the package and returns it based on the path
        /// </summary>
        /// <param name="path">Path of the package</param>
        /// <param name="mode">File Mode of the package</param>
        /// <param name="access">File Access</param>
        /// <returns>Package</returns>
        public static PnPPackage Open(string path, FileMode mode, FileAccess access)
        {
            PnPPackage package = new PnPPackage();
            package.Package = Package.Open(path, mode, access);
            package.EnsureMandatoryPackageComponents();
            return package;
        }

        /// <summary>
        /// Opens the package and returns it based on the stream
        /// </summary>
        /// <param name="stream">The stream</param>
        /// <param name="mode">File Mode of the package</param>
        /// <param name="access">File Access</param>
        /// <returns>Package</returns>
        public static PnPPackage Open(Stream stream, FileMode mode, FileAccess access)
        {
            PnPPackage package = new PnPPackage();
            package.Package = Package.Open(stream, mode, access);
            package.EnsureMandatoryPackageComponents();
            return package;
        }

        /// <summary>
        /// Adds file to the package
        /// </summary>
        /// <param name="fileName">Name of the file</param>
        /// <param name="value">Value of the file</param>
        public void AddFile(string fileName, Byte[] value)
        {
            fileName = fileName.TrimStart('/');
            string uriStr = U_DIR_FILES + fileName;
            PackagePart part = CreatePackagePart(R_PROVISIONINGTEMPLATE_FILE, CT_FILE, uriStr, FilesOriginPart);
            SetPackagePartValue(value, part);
        }

        /// <summary>
        /// Clear the files having package parts with specific relationship type
        /// </summary>
        public void ClearFiles()
        {
            ClearPackagePartsWithRelationshipType(R_PROVISIONINGTEMPLATE_FILE, FilesOriginPart);
        }

#endregion

#region Package Helper Methods

        private void EnsureMandatoryPackageComponents()
        {
            // Manifest
            EnsurePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_MANIFEST, CT_PROVISIONINGTEMPLATE_MANIFEST, U_PROVISIONINGTEMPLATE_MANIFEST, null);

            // Properties
            EnsurePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_PROPERTIES, CT_PROVISIONINGTEMPLATE_PROPERTIES, U_PROVISIONINGTEMPLATE_PROPERTIES, ManifestPart);

            // Files origin
            EnsurePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_FILES_ORIGIN, CT_ORIGIN, U_FILES_ORIGIN, ManifestPart);

            // Files map
            EnsurePackagePartWithRelationshipType(R_PROVISIONINGTEMPLATE_FILES_MAP, CT_PROVISIONINGTEMPLATE_FILES_MAP, U_PROVISIONINGTEMPLATE_FILES_MAP, ManifestPart);
        }

        private PackagePart EnsurePackagePartWithRelationshipType(string relType, string contentType, string uriStr, PackagePart parent)
        {
            PackagePart part = GetSinglePackagePartWithRelationshipType(relType, parent);
            if (part == null)
            {
                part = CreatePackagePart(relType, contentType, uriStr, parent);
            }
            return part;
        }

        private PackagePart EnsurePackagePartWithUri(string relType, string contentType, string uriStr, PackagePart parent)
        {
            Uri partUri = new Uri(uriStr, UriKind.Relative);
            PackagePart part = null;
            try
            {
                part = Package.GetPart(partUri);
            }
            catch { }
            if (part == null)
            {
                part = CreatePackagePart(relType, contentType, uriStr, parent);
            }
            return part;
        }

        private PackagePart GetSinglePackagePartWithRelationshipType(string relType, PackagePart parent)
        {
            PackageRelationshipCollection rels = parent == null ? Package.GetRelationshipsByType(relType) : parent.GetRelationshipsByType(relType);
            PackageRelationship rel = null;
            foreach (PackageRelationship r in rels)
            {
                if (r.TargetMode == TargetMode.Internal)
                {
                    rel = r;
                    break;
                }
            }
            if (rel != null)
            {
                return Package.GetPart(rel.TargetUri);
            }
            return null;
        }

        private List<PackagePart> GetAllPackagePartsWithRelationshipType(string relType, PackagePart parent)
        {
            PackageRelationshipCollection rels = parent == null ? Package.GetRelationshipsByType(relType) : parent.GetRelationshipsByType(relType);
            List<PackagePart> pkgList = new List<PackagePart>();
            foreach (PackageRelationship rel in rels)
            {
                if (rel.TargetMode == TargetMode.Internal)
                    pkgList.Add(Package.GetPart(rel.TargetUri));
            }
            return pkgList;
        }

        private T GetXamlSerializedPackagePartValue<T>(PackagePart part) where T : class
        {
            if (part == null)
                return null;

            T obj = null;
            using (Stream stream = part.GetStream(FileMode.Open))
            {
                if (stream.Length == 0)
                    return null;
                obj = (T)XamlServices.Load(stream);
            }
            return obj;
        }

        private void SetXamlSerializedPackagePartValue(object value, PackagePart part)
        {
            if (value == null)
                return;

            using (Stream stream = part.GetStream(FileMode.Create))
            {
                string partStr = XamlServices.Save(value);
                using (stream)
                {
                    using (StreamWriter writer = new StreamWriter(stream))
                    {
                        writer.Write(partStr);
                    }
                }
            }
        }

        private void SetPackagePartValue(Byte[] value, PackagePart part)
        {
            if (value == null)
            {
                value = new byte[] { };
            }

            using (Stream stream = part.GetStream(FileMode.Create))
            {
                using (stream)
                {
                    stream.Write(value, 0, value.Length);
                }
            }
        }

        private PackagePart CreatePackagePart(string relType, string contentType, string uriStr, PackagePart parent)
        {
            // create part & relationship
            Uri uri = GetUri(uriStr);
            PackagePart part = Package.CreatePart(uri, contentType, PACKAGE_COMPRESSION_LEVEL);
            if (parent == null)
            {
                Package.CreateRelationship(uri, TargetMode.Internal, relType);
            }
            else
            {
                parent.CreateRelationship(uri, TargetMode.Internal, relType);
            }
            return part;
        }

        private void ClearPackagePartsWithRelationshipType(string relType, PackagePart parent)
        {
            ClearPackagePartsWithRelationshipType(relType, parent, null);
        }

        private void ClearPackagePartsWithRelationshipType(string relType, PackagePart parent, string partUri)
        {
            PackageRelationshipCollection rels = parent == null ? Package.GetRelationshipsByType(relType) : parent.GetRelationshipsByType(relType);
            List<string> relIds = new List<string>();
            foreach (PackageRelationship r in rels)
            {
                if (r.TargetMode == TargetMode.Internal)
                {
                    if (partUri == null || r.TargetUri.ToString() == partUri)
                    {
                        Package.DeletePart(r.TargetUri);
                        relIds.Add(r.Id);
                    }
                }
            }

            foreach (string relId in relIds)
            {
                if (parent == null)
                {
                    Package.DeleteRelationship(relId);
                }
                else
                {
                    parent.DeleteRelationship(relId);
                }
            }
        }

        private List<byte[]> ReadPackagePartListBytes(List<PackagePart> partList)
        {
            List<byte[]> result = new List<byte[]>();
            if (partList != null)
            {
                foreach (PackagePart p in partList)
                {
                    result.Add(ReadPackagePartBytes(p));
                }
            }
            return result;
        }

        private Byte[] ReadPackagePartBytes(PackagePart part)
        {
            if (part == null)
                return null;

            Byte[] bytes;
            using (Stream stream = part.GetStream())
            {
                long size = stream.Length;

                //TODO: fix method to support long streams
                if (size > Int32.MaxValue)
                    throw new ArgumentOutOfRangeException(nameof(part),"Long streams are not supported.");

                bytes = new byte[size];
                stream.Read(bytes, 0, (int)size);
            }
            return bytes;

        }

#endregion

#region Generic Helper methods

        private Uri GetUri(string uriStr)
        {
            return new Uri(uriStr, UriKind.Relative);
        }

#endregion

        void IDisposable.Dispose()
        {
            if (Package != null)
                Package.Close();
        }
    }
}
