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

        public override int GetHashCode()
        {
            return (String.Format("{0}",
                (this.Name != null ? this.Name.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is User))
            {
                return (false);
            }
            return (Equals((User)obj));
        }

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
