using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Drive
{
    /// <summary>
    /// Defines a DriveFolder object
    /// </summary>
    public partial class DriveFolder : DriveFolderBase
    {
        #region Public Members

        /// <summary>
        /// Defines the Name of the Folder in OneDrive for Business
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Defines the Source path of the folder in OneDrive for Business
        /// </summary>
        public string Src { get; set; }

        /// <summary>
        /// The Overwrite flag for the File items in the Directory
        /// </summary>
        public bool Overwrite { get; set; }

        /// <summary>
        /// Defines whether to recursively browse through all the child folders of the source Folder
        /// </summary>
        public bool Recursive { get; set; }

        /// <summary>
        /// The file Extensions (lower case) to include while uploading the source Folder
        /// </summary>
        public string IncludedExtensions { get; set; }

        /// <summary>
        /// The file Extensions (lower case) to exclude while uploading the source Folder
        /// </summary>
        public string ExcludedExtensions { get; set; }

        #endregion

        #region Constructors

        public DriveFolder() : base()
        {
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        protected override int GetInheritedHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|",
                this.Name?.GetHashCode() ?? 0,
                this.Src?.GetHashCode() ?? 0,
                this.Overwrite.GetHashCode(),
                this.Recursive.GetHashCode(),
                this.IncludedExtensions?.GetHashCode() ?? 0,
                this.ExcludedExtensions?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        protected override bool EqualsInherited(DriveFolderBase other)
        {
            if (!(other is DriveFolder otherTyped))
            {
                return (false);
            }

            return (this.Name == otherTyped.Name &&
                this.Src == otherTyped.Src &&
                this.Overwrite == otherTyped.Overwrite &&
                this.Recursive == otherTyped.Recursive &&
                this.IncludedExtensions == otherTyped.IncludedExtensions &&
                this.ExcludedExtensions == otherTyped.ExcludedExtensions
                );
        }

        #endregion
    }
}
