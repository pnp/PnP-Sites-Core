using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class PropertyBagEntry : BaseModel, IEquatable<PropertyBagEntry>
    {
        #region Properties
        /// <summary>
        /// Gets or sets the Key for property bag entry
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Gets or sets the Value for the property bag entry
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Gets or sets the Indexed flag for property bag entry
        /// </summary>
        public bool Indexed { get; set; }

        /// <summary>
        /// Gets or sets the Overwrite flag for property bag entry
        /// </summary>
        public bool Overwrite { get; set; }
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
                this.Indexed.GetHashCode(),
                this.Overwrite.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with PropertyBagEntry
        /// </summary>
        /// <param name="obj">Object</param>
        /// <returns>true if the current object is equal to the PropertyBagEntry</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is PropertyBagEntry))
            {
                return (false);
            }
            return (Equals((PropertyBagEntry)obj));
        }

        /// <summary>
        /// Compares PropertBag object based on Key, Value, Indexed and Overwrite properties.
        /// </summary>
        /// <param name="other">PropertyBagEntry object</param>
        /// <returns>true if the PropertyBagEntry object is equal to the current object; otherwise, false.</returns>
        public bool Equals(PropertyBagEntry other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Key == other.Key &&
                this.Value == other.Value &&
                this.Indexed == other.Indexed &&
                this.Overwrite == other.Overwrite);
        }

        #endregion
    }
}
