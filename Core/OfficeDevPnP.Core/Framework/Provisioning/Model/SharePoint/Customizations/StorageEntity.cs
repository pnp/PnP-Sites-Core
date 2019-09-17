using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a StorageEntity to provision
    /// </summary>
    public partial class StorageEntity : BaseModel, IEquatable<StorageEntity>
    {
        #region Properties
        /// <summary>
        /// Gets or sets the Key for the StorageEntity
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Gets or sets the Value for the StorageEntity
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Gets or sets the Comment flag for the StorageEntity
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// Gets or sets the Description flag for the StorageEntity
        /// </summary>
        public string Description { get; set; }
        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                (this.Key != null ? this.Key.GetHashCode() : 0),
                (this.Value != null ? this.Value.GetHashCode() : 0),
                (this.Comment != null ? this.Comment.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with StorageEntity
        /// </summary>
        /// <param name="obj">Object</param>
        /// <returns>true if the current object is equal to the StorageEntity</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is StorageEntity))
            {
                return (false);
            }
            return (Equals((StorageEntity)obj));
        }

        /// <summary>
        /// Compares StorageEntity object based on Key, Value, Comment and Description properties.
        /// </summary>
        /// <param name="other">StorageEntity object</param>
        /// <returns>true if the StorageEntity object is equal to the current object; otherwise, false.</returns>
        public bool Equals(StorageEntity other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Key == other.Key &&
                this.Value == other.Value &&
                this.Comment == other.Comment &&
                this.Description == other.Description);
        }

        #endregion
    }
}