using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that defines a User or group in the provisioning template
    /// </summary>
    public partial class User : BaseModel, IEquatable<User>
    {
        #region Public Members

        /// <summary>
        /// The User email Address or the group name.
        /// </summary>
        public string Name { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}",
                (this.Name != null ? this.Name.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with User
        /// </summary>
        /// <param name="obj">Object that represents User</param>
        /// <returns>true if the current object is equal to the User</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is User))
            {
                return (false);
            }
            return (Equals((User)obj));
        }

        /// <summary>
        /// Compares User object based on Name
        /// </summary>
        /// <param name="other">User object</param>
        /// <returns>true if the User object is equal to the current object; otherwise, false.</returns>
        public bool Equals(User other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Name == other.Name);
        }

        #endregion
    }
}
