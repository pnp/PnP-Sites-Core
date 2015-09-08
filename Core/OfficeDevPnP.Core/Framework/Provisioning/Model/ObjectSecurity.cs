using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class ObjectSecurity : IEquatable<ObjectSecurity>
    {
        #region Private Members

        private List<RoleAssignment> _roleAssignments = new List<RoleAssignment>();

        #endregion

        #region Constructors

        public ObjectSecurity() { }

        public ObjectSecurity(IEnumerable<RoleAssignment> roleAssignments)
        {
            if (roleAssignments != null)
            {
                this._roleAssignments.AddRange(roleAssignments);
            }
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Role Assignments for a target Principal
        /// </summary>
        public List<RoleAssignment> RoleAssignments
        {
            get { return this._roleAssignments; }
            private set { this._roleAssignments = value; }
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
                this.RoleAssignments.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
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
                this.RoleAssignments.DeepEquals(other.RoleAssignments) &&
                this.CopyRoleAssignments == other.CopyRoleAssignments &&
                this.ClearSubscopes == other.ClearSubscopes
                );
        }

        #endregion
    }
}
