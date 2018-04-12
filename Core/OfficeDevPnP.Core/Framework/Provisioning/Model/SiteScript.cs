using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain object to define a SiteScript to provision
    /// </summary>
    public partial class SiteScript : BaseModel, IEquatable<SiteScript>
    {
        #region Properties

        /// <summary>
        /// Gets or sets the Title for the SiteScript
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Description flag for the SiteScript
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the JsonFilePath flag for the SiteScript
        /// </summary>
        public string JsonFilePath { get; set; }

        /// <summary>
        /// Defines whether to overwrite the SiteDesign or not
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
            return (String.Format("{0}|{1}|{2}|{3}|",
                (this.Title != null ? this.Title.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.JsonFilePath != null ? this.JsonFilePath.GetHashCode() : 0),
                this.Overwrite.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteScript
        /// </summary>
        /// <param name="obj">Object</param>
        /// <returns>true if the current object is equal to the SiteScript</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteScript))
            {
                return (false);
            }
            return (Equals((SiteScript)obj));
        }

        /// <summary>
        /// Compares SiteScript object based on Title, Description, and JsonFilePath properties.
        /// </summary>
        /// <param name="other">SiteScript object</param>
        /// <returns>true if the SiteScript object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteScript other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Title == other.Title &&
                this.Description == other.Description &&
                this.JsonFilePath == other.JsonFilePath &&
                this.Overwrite == other.Overwrite
                );
        }

        #endregion
    }
}
