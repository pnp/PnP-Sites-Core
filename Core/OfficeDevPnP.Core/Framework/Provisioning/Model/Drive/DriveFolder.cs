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
    public partial class DriveFolder : BaseModel, IEquatable<DriveFolder>
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
        /// Defines a collection of DriveFolder items
        /// </summary>
        public DriveFolderCollection DriveFolders { get; private set; }

        /// <summary>
        /// Defines a collection of DriveFile items
        /// </summary>
        public DriveFileCollection DriveFiles { get; private set; }

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
            this.DriveFolders = new DriveFolderCollection(this.ParentTemplate);
            this.DriveFiles = new DriveFileCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|",
                DriveFolders.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                DriveFiles.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Name?.GetHashCode() ?? 0,
                this.Src?.GetHashCode() ?? 0,
                this.Overwrite.GetHashCode(),
                this.Recursive.GetHashCode(),
                this.IncludedExtensions?.GetHashCode() ?? 0,
                this.ExcludedExtensions?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with DriveFolder class
        /// </summary>
        /// <param name="obj">Object that represents DriveFolder</param>
        /// <returns>Checks whether object is DriveFolder class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DriveFolder))
            {
                return (false);
            }
            return (Equals((DriveFolder)obj));
        }

        /// <summary>
        /// Compares DriveFolder object based on Drivefolders, DriveFiles, Name, Src,
        /// Overwrite, Recursive, IncludedExtensions, and ExcludedExtensions
        /// </summary>
        /// <param name="other">DriveFolder Class object</param>
        /// <returns>true if the DriveFolder object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DriveFolder other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.DriveFolders.DeepEquals(other.DriveFolders) &&
                this.DriveFiles.DeepEquals(other.DriveFiles) &&
                this.Name == other.Name &&
                this.Src == other.Src &&
                this.Overwrite == other.Overwrite &&
                this.Recursive == other.Recursive &&
                this.IncludedExtensions == other.IncludedExtensions &&
                this.ExcludedExtensions == other.ExcludedExtensions
                );
        }

        #endregion
    }
}
