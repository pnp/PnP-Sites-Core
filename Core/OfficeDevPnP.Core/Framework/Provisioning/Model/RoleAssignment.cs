using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Role Assignment for a target Principal
    /// </summary>
    public partial class RoleAssignment : BaseModel, IEquatable<RoleAssignment>
    {
        #region Public Members

        /// <summary>
        /// Defines the Role to which the assignment will apply
        /// </summary>
        public String Principal { get; set; }

        /// <summary>
        /// Defines the Role to which the assignment will apply
        /// </summary>
        public String RoleDefinition { get; set; }

        /// <summary>
        /// Allows to remove a role assignment, instead of adding it. It is an optional attribute, and by default it assumes a value of false.
        /// </summary>
        public Boolean Remove { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                (this.Principal != null ? this.Principal.GetHashCode() : 0),
                (this.RoleDefinition != null ? this.RoleDefinition.GetHashCode() : 0),
                this.Remove.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with RoleAssignment
        /// </summary>
        /// <param name="obj">Object that represents RoleAssignment</param>
        /// <returns>true if the current object is equal to the RoleAssignment</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is RoleAssignment))
            {
                return (false);
            }
            return (Equals((RoleAssignment)obj));
        }

        /// <summary>
        /// Compares RoleAssignment object based on Principal and RoleDefinition
        /// </summary>
        /// <param name="other">RoleAssignment object</param>
        /// <returns>true if the RoleAssignment object is equal to the current object; otherwise, false.</returns>
        public bool Equals(RoleAssignment other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Principal == other.Principal &&
                this.RoleDefinition == other.RoleDefinition &&
                this.Remove == other.Remove);
        }

        #endregion
    }
}
