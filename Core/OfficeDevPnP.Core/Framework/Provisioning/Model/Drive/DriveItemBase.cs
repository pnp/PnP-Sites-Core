using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Drive
{
    /// <summary>
    /// Defines a DriveItem object
    /// </summary>
    public abstract partial class DriveItemBase : BaseModel, IEquatable<DriveItemBase>
    {
        #region Public members

        /// <summary>
        /// Defines a collection of DriveFolder items
        /// </summary>
        public DriveFolderCollection DriveFolders { get; private set; }

        /// <summary>
        /// Defines a collection of DriveFile items
        /// </summary>
        public DriveFileCollection DriveFiles { get; private set; }

        #endregion

        #region Constructors

        public DriveItemBase() : base()
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
            return (String.Format("{0}|{1}|",
                DriveFolders.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                DriveFiles.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with DriveItemBase class
        /// </summary>
        /// <param name="obj">Object that represents DriveItemBase</param>
        /// <returns>Checks whether object is DriveItemBase class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DriveItemBase))
            {
                return (false);
            }
            return (Equals((DriveItemBase)obj));
        }

        /// <summary>
        /// Compares DriveItemBase object based on Drivefolders, and DriveFiles
        /// </summary>
        /// <param name="other">DriveItemBase Class object</param>
        /// <returns>true if the DriveItemBase object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DriveItemBase other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.DriveFolders.DeepEquals(other.DriveFolders) &&
                this.DriveFiles.DeepEquals(other.DriveFiles)
                );
        }

        #endregion
    }
}
