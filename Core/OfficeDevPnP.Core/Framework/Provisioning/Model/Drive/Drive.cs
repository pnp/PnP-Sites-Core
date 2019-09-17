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
    public partial class Drive : BaseModel, IEquatable<Drive>
    {
        #region Public members

        /// <summary>
        /// Defines a collection of DriveRoot items
        /// </summary>
        public DriveRootCollection DriveRoots { get; private set; }

        #endregion

        #region Constructors

        public Drive() : base()
        {
            this.DriveRoots = new DriveRootCollection(this.ParentTemplate);
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}",
                DriveRoots.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Drive class
        /// </summary>
        /// <param name="obj">Object that represents Drive</param>
        /// <returns>Checks whether object is Drive class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Drive))
            {
                return (false);
            }
            return (Equals((Drive)obj));
        }

        /// <summary>
        /// Compares Drive object based on DriveRoots
        /// </summary>
        /// <param name="other">Drive Class object</param>
        /// <returns>true if the Drive object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Drive other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.DriveRoots.DeepEquals(other.DriveRoots)
                );
        }

        #endregion
    }
}
