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

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                (this.Principal != null ? this.Principal.GetHashCode() : 0),
                (this.RoleDefinition != null ? this.RoleDefinition.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is RoleAssignment))
            {
                return (false);
            }
            return (Equals((RoleAssignment)obj));
        }

        public bool Equals(RoleAssignment other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Principal == other.Principal &&
                this.RoleDefinition == other.RoleDefinition);
        }

        #endregion
    }
}
