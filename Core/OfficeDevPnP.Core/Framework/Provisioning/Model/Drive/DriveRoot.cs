using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Drive
{
    /// <summary>
    /// Defines a Drive object
    /// </summary>
    public partial class DriveRoot : BaseModel, IEquatable<DriveRoot>
    {
        #region Public members

        /// <summary>
        /// The DriveUrl of the target DriveRoot
        /// </summary>
        public String DriveUrl { get; set; }

        /// <summary>
        /// Defines the RootFolder of a DriveRoot item
        /// </summary>
        public DriveRootFolder RootFolder { get; set; }

        #endregion

        #region Constructors

        public DriveRoot() : base()
        {
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
                DriveUrl?.GetHashCode() ?? 0,
                RootFolder.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with DriveRoot class
        /// </summary>
        /// <param name="obj">Object that represents DriveRoot</param>
        /// <returns>Checks whether object is DriveRoot class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DriveRoot))
            {
                return (false);
            }
            return (Equals((DriveRoot)obj));
        }

        /// <summary>
        /// Compares DriveRoot object based on DriveUrl, and RootFolder
        /// </summary>
        /// <param name="other">User DriveRoot object</param>
        /// <returns>true if the DriveRoot object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DriveRoot other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.DriveUrl == other.DriveUrl &&
                this.RootFolder.Equals(other.RootFolder)
                );
        }

        #endregion
    }
}
