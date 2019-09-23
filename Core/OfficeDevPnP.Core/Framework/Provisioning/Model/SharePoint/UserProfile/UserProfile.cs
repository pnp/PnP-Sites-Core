using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.SPUPS
{
    /// <summary>
    /// Defines a UserProfile object for SharePoint
    /// </summary>
    public partial class UserProfile : BaseModel, IEquatable<UserProfile>
    {
        #region Public members

        /// <summary>
        /// Properties of the UserProfile
        /// </summary>
        public Dictionary<string, string> Properties { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// The Target User of the target UserProfile
        /// </summary>
        public String TargetUser { get; set; }

        /// <summary>
        /// The Targe Group of the target UserProfile
        /// </summary>
        public String TargetGroup { get; set; }

        #endregion

        #region Constructors

        public UserProfile() : base()
        {
        }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                TargetUser?.GetHashCode() ?? 0,
                TargetGroup?.GetHashCode() ?? 0,
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with UserProfile class
        /// </summary>
        /// <param name="obj">Object that represents UserProfile</param>
        /// <returns>Checks whether object is UserProfile class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is UserProfile))
            {
                return (false);
            }
            return (Equals((UserProfile)obj));
        }

        /// <summary>
        /// Compares UserProfile object based on TargetUser, TargetGroup, and Properties
        /// </summary>
        /// <param name="other">User UserProfile object</param>
        /// <returns>true if the UserProfile object is equal to the current object; otherwise, false.</returns>
        public bool Equals(UserProfile other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.TargetUser == other.TargetUser &&
                this.TargetGroup == other.TargetGroup &&
                this.Properties.DeepEquals(other.Properties)
                );
        }

        #endregion
    }
}
