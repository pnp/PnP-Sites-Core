using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines an available Package for the AppCatalog.
    /// </summary>
    public partial class Package : BaseModel, IEquatable<Package>
    {
        #region Public Members

        /// <summary>
        /// Defines the Package Id of the Package, optional attribute.
        /// </summary>
        public String PackageId { get; set; }

        /// <summary>
        /// Defines the Src of the Package, optional attribute.
        /// </summary>
        public String Src { get; set; }

        /// <summary>
        /// Defines the Action to execute with the Package in the AppCatalog, required attribute.
        /// </summary>
        public PackageAction Action { get; set; }

        /// <summary>
        /// Defines whether to skip the feature deployment for tenant-wide enabled packages
        /// </summary>
        public Boolean SkipFeatureDeployment { get; set; }

        /// <summary>
        /// Defines whether to overwrite an already existing package in the AppCatalog
        /// </summary>
        public Boolean Overwrite { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.PackageId?.GetHashCode() ?? 0,
                this.Src?.GetHashCode() ?? 0,
                this.Action.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Package class
        /// </summary>
        /// <param name="obj">Object that represents Package</param>
        /// <returns>Checks whether object is Package class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Package))
            {
                return (false);
            }
            return (Equals((Package)obj));
        }

        /// <summary>
        /// Compares Package object based on PackageId, Src, Action
        /// </summary>
        /// <param name="other">Package Class object</param>
        /// <returns>true if the Package object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Package other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.PackageId== other.PackageId &&
                this.Src == other.Src &&
                this.Action == other.Action
                );
        }

        #endregion
    }

    /// <summary>
    /// Defines the actions for a Package in the AppCatalog
    /// </summary>
    public enum PackageAction
    {
        /// <summary>
        /// Instructs the engine to upload the package in the App Catalog. The Src is required.
        /// </summary>
        Upload,
        /// <summary>
        /// Instructs the engine to publish the package in the App Catalog. The PackageId is required.
        /// </summary>
        Publish,
        /// <summary>
        /// Instructs the engine to upload and publish the package in the App Catalog. The PackageId is required.
        /// </summary>
        UploadAndPublish,
        /// <summary>
        /// Instructs the engine to remove the package from the App Catalog. The PackageId is required.
        /// </summary>
        Remove,
    }
}
