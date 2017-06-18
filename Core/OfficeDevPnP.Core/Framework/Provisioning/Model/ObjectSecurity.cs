using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class ObjectSecurity : BaseModel, IEquatable<ObjectSecurity>
    {
        #region Private Members

        private RoleAssignmentCollection _roleAssignments;

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for ObjectSecurity class
        /// </summary>
        public ObjectSecurity()
        {
            this._roleAssignments = new RoleAssignmentCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for ObjectSecurity class
        /// </summary>
        /// <param name="roleAssignments">RoleAssignments for security component</param>
        public ObjectSecurity(IEnumerable<RoleAssignment> roleAssignments):
            this()
        {
            this.RoleAssignments.AddRange(roleAssignments);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Role Assignments for a target Principal
        /// </summary>
        public RoleAssignmentCollection RoleAssignments
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
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.RoleAssignments.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.CopyRoleAssignments.GetHashCode(),
                this.ClearSubscopes.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ObjectSecurity
        /// </summary>
        /// <param name="obj">Object that represents ObjectSecurity</param>
        /// <returns>true if the current object is equal to the ObjectSecurity</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ObjectSecurity))
            {
                return (false);
            }
            return (Equals((ObjectSecurity)obj));
        }

        /// <summary>
        /// Compares ObjectSecurity object based on RoleAssignments, CopyRoleAssignments and ClearSubscopes properties.
        /// </summary>
        /// <param name="other">ObjectSecurity object</param>
        /// <returns>true if the ObjectSecurity object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ObjectSecurity other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.RoleAssignments.DeepEquals(other.RoleAssignments) &&
                this.CopyRoleAssignments == other.CopyRoleAssignments &&
                this.ClearSubscopes == other.ClearSubscopes
                );
        }

        #endregion
    }
}
