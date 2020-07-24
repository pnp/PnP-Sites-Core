using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines a user for a the Team
    /// </summary>
    public partial class TeamSecurityUser : BaseModel, IEquatable<TeamSecurityUser>
    {
        #region Public Members

        /// <summary>
        /// Defines User Principal Name (UPN) of the target user
        /// </summary>
        public String UserPrincipalName { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                UserPrincipalName?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamSecurityUser class
        /// </summary>
        /// <param name="obj">Object that represents TeamSecurityUser</param>
        /// <returns>Checks whether object is TeamSecurityUser class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamSecurityUser))
            {
                return (false);
            }
            return (Equals((TeamSecurityUser)obj));
        }

        /// <summary>
        /// Compares TeamSecurityUser object based on UserPrincipalName
        /// </summary>
        /// <param name="other">TeamSecurityUser Class object</param>
        /// <returns>true if the TeamSecurityUser object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamSecurityUser other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.UserPrincipalName == other.UserPrincipalName);
        }

        #endregion
    }
}
