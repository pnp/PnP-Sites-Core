using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML
{
    /// <summary>
    /// Defines a PnP OpenXML package file
    /// </summary>
    public partial class PnPPackage
    {
        [Obsolete("Folders are now handled inside the mapping file. This method will be removed in the September 2016 release.")]
        public void AddFile(string fileName, string folder, Byte[] value)
        {
            fileName = fileName.TrimStart('/');
            folder = !String.IsNullOrEmpty(folder) ? (folder.TrimStart('/').TrimEnd('/') + "/") : String.Empty;
            string uriStr = U_DIR_FILES + folder + fileName;
            PackagePart part = CreatePackagePart(R_PROVISIONINGTEMPLATE_FILE, CT_FILE, uriStr, FilesOriginPart);
            SetPackagePartValue(value, part);
        }
    }
}
