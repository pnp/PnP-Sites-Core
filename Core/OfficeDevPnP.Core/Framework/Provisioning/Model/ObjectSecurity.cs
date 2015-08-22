using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class ObjectSecurity : IEquatable<ObjectSecurity>
    {
        #region Private Members

        private RoleAssignment _roleAssignment = new RoleAssignment();

        #endregion

        #region Public Members

        /// <summary>
        /// Role Assignment for a target Principal
        /// </summary>
        public RoleAssignment RoleAssignment
        {
            get { return this._roleAssignment; }
            set { this._roleAssignment = value; }
        }

        /// <summary>
        /// Defines whether to copy role assignments or not while breaking role inheritance
        /// </summary>
        public Boolean CopyRoleAssignments { get; set; }

        /// <summary>
        /// Defines whether to clear subscopes or not while breaking role inheritance
        /// </summary>
        public Boolean ClearSubscopes { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.RoleAssignment.GetHashCode(),
                this.CopyRoleAssignments.GetHashCode(),
                this.ClearSubscopes.GetHashCode()
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ObjectSecurity))
            {
                return (false);
            }
            return (Equals((ObjectSecurity)obj));
        }

        public bool Equals(ObjectSecurity other)
        {
            return (
                this.RoleAssignment == other.RoleAssignment &&
                this.CopyRoleAssignments == other.CopyRoleAssignments &&
                this.ClearSubscopes == other.ClearSubscopes
                );
        }

        #endregion
    }
}
