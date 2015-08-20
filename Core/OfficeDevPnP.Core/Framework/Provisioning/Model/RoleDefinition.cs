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
        /// <summary>
        /// Defines the Permissions of the Role Definition
        /// </summary>
        public List<Microsoft.SharePoint.Client.PermissionKind> Permissions { get; set; }

        /// <summary>
        /// Defines the Name of the Role Definition
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Defines the Description of the Role Definition
        /// </summary>
        public String Description { get; set; }

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|",
                this.Permissions,
                this.Name,
                this.Description
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
            return (this.Permissions.DeepEquals(other.Permissions) &&
                this.Name == other.Name &&
                this.Description == other.Description );
        }

        #endregion
    }
}
