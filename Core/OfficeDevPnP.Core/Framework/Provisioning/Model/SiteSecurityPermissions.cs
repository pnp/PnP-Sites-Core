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
        #region Private Members

        private List<RoleDefinition> _roleDefinitions = new List<RoleDefinition>();
        private List<RoleAssignment> _roleAssignments = new List<RoleAssignment>();

        #endregion

        #region Public Members

        /// <summary>
        /// List of Role Definitions for the Site
        /// </summary>
        public List<RoleDefinition> RoleDefinitions
        {
            get { return this._roleDefinitions; }
            private set { this._roleDefinitions = value; }
        }

        /// <summary>
        /// List of Role Assignments for the Site
        /// </summary>
        public List<RoleAssignment> RoleAssignments
        {
            get { return this._roleAssignments; }
            private set { this._roleAssignments = value; }
        }

        #endregion
    }
}
