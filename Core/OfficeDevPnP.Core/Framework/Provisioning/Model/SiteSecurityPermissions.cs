using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Permission settings for the target Site
    /// </summary>
    public class SiteSecurityPermissions
    {
        /// <summary>
        /// List of Role Definitions for the Site
        /// </summary>
        public List<RoleDefinition> RoleDefinitions { get; set; }

        /// <summary>
        /// List of Role Assignments for the Site
        /// </summary>
        public List<RoleAssignment> RoleAssignments { get; set; }
    }
}
