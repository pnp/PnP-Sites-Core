using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Drive
{
    /// <summary>
    /// Defines a DriveFolderBase object
    /// </summary>
    public abstract partial class DriveFolderBase : BaseModel, IEquatable<DriveFolderBase>
    {
        #region Public Members

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

        public DriveFolderBase() : base()
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
            return (String.Format("{0}|{1}|{2}|",
                DriveFolders.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                DriveFiles.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.GetInheritedHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Returns the HashCode of the members of any inherited type
        /// </summary>
        /// <returns></returns>
        protected abstract int GetInheritedHashCode();

        /// <summary>
        /// Compares object with DriveFolderBase class
        /// </summary>
        /// <param name="obj">Object that represents DriveFolderBase</param>
        /// <returns>Checks whether object is DriveFolderBase class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DriveFolderBase))
            {
                return (false);
            }
            return (Equals((DriveFolderBase)obj));
        }

        /// <summary>
        /// Compares DriveFolderBase object based on Drivefolders, DriveFiles
        /// </summary>
        /// <param name="other">DriveFolderBase Class object</param>
        /// <returns>true if the DriveFolderBase object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DriveFolderBase other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.DriveFolders.DeepEquals(other.DriveFolders) &&
                this.DriveFiles.DeepEquals(other.DriveFiles) &&
                this.EqualsInherited(other)
                );
        }

        /// <summary>
        /// Compares the HashCode of the members of any inherited type
        /// </summary>
        /// <returns></returns>
        protected abstract bool EqualsInherited(DriveFolderBase other);

        #endregion
    }
}
