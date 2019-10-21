using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines an Add-in to provision
    /// </summary>
    public partial class AddIn : BaseModel, IEquatable<AddIn>
    {
        #region Public Members

        /// <summary>
        /// Defines the .app file of the SharePoint Add-in to provision
        /// </summary>
        public String PackagePath { get; set; }

        /// <summary>
        /// Defines the Source of the SharePoint Add-in to provision
        /// </summary>
        /// <remarks>
        /// Possible values are: CorporateCatalog, DeveloperSite, InvalidSource, Marketplace, ObjectModel, RemoteObjectModel
        /// </remarks>
        public String Source { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                PackagePath?.GetHashCode() ?? 0,
                Source?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with AddIn class
        /// </summary>
        /// <param name="obj">Object that represents AddIn</param>
        /// <returns>Checks whether object is AddIn class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is AddIn))
            {
                return (false);
            }
            return (Equals((AddIn)obj));
        }

        /// <summary>
        /// Compares AddIn object based on PackagePath and source
        /// </summary>
        /// <param name="other">AddIn Class object</param>
        /// <returns>true if the AddIn object is equal to the current object; otherwise, false.</returns>
        public bool Equals(AddIn other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.PackagePath == other.PackagePath &&
                this.Source == other.Source
                );
        }

        #endregion
    }
}
