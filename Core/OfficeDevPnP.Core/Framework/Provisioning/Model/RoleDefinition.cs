using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class RoleDefinition : IEquatable<RoleDefinition>
    {
        #region Private Members

        private List<Microsoft.SharePoint.Client.PermissionKind> _permissions = new List<Microsoft.SharePoint.Client.PermissionKind>();

        #endregion

        #region Constructors

        public RoleDefinition() { }

        public RoleDefinition(IEnumerable<Microsoft.SharePoint.Client.PermissionKind> permissions)
        {
            if (permissions != null)
            {
                this._permissions.AddRange(permissions);
            }
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the Permissions of the Role Definition
        /// </summary>
        public List<Microsoft.SharePoint.Client.PermissionKind> Permissions
        {
            get { return this._permissions; }
            private set { this._permissions = value; }
        }

        /// <summary>
        /// Defines the Name of the Role Definition
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Defines the Description of the Role Definition
        /// </summary>
        public String Description { get; set; }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.Permissions.GetHashCode(),
                this.Name.GetHashCode(),
                this.Description.GetHashCode(),
                this.Permissions.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is RoleDefinition))
            {
                return (false);
            }
            return (Equals((RoleDefinition)obj));
        }

        public bool Equals(RoleDefinition other)
        {
            return (this.Name == other.Name &&
                this.Description == other.Description &&
                this.Permissions.DeepEquals(other.Permissions));
        }

        #endregion
    }
}
