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
    public partial class DriveRootFolder : DriveFolderBase
    {
        #region Public Members

        #endregion

        #region Constructors

        public DriveRootFolder() : base()
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
            return (0);
        }

        protected override bool EqualsInherited(DriveFolderBase other)
        {
            if (!(other is DriveRootFolder otherTyped))
            {
                return (false);
            }

            return (true);
        }

        #endregion
    }
}
