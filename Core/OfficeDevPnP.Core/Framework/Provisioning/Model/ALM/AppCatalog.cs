using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the AppCatalog packages to provision
    /// </summary>
    public partial class AppCatalog : BaseModel, IEquatable<AppCatalog>
    {
        #region Private Members

        private PackageCollection _packages;

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public AppCatalog()
        {
            this._packages = new Model.PackageCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Custom constructor
        /// </summary>
        public AppCatalog(PackageCollection packages): base()
        {
            this._packages.AddRange(packages);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the AppCatalog settings to provision
        /// </summary>
        public PackageCollection Packages
        {
            get { return this._packages; }
            private set { this._packages = value; }
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                this.Packages.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with AppCatalog
        /// </summary>
        /// <param name="obj">Object that represents AppCatalog</param>
        /// <returns>true if the current object is equal to the AppCatalog</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is AppCatalog))
            {
                return (false);
            }
            return (Equals((AppCatalog)obj));
        }

        /// <summary>
        /// Compares AppCatalog object based on Packages properties.
        /// </summary>
        /// <param name="other">AppCatalog object</param>
        /// <returns>true if the AppCatalog object is equal to the current object; otherwise, false.</returns>
        public bool Equals(AppCatalog other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Packages.DeepEquals(other.Packages)
                );
        }

        #endregion
    }
}
