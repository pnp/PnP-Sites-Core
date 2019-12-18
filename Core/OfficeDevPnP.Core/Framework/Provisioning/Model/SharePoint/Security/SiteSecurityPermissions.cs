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
    public partial class SiteSecurityPermissions : BaseModel
    {
        #region Private Members

        private RoleDefinitionCollection _roleDefinitions;
        private RoleAssignmentCollection _roleAssignments;

        #endregion

        #region Constructor
        /// <summary>
        /// Constructor for SiteSecurityPermissions class
        /// </summary>
        public SiteSecurityPermissions()
        {
            this._roleDefinitions = new RoleDefinitionCollection(this.ParentTemplate);
            this._roleAssignments = new RoleAssignmentCollection(this.ParentTemplate);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// List of Role Definitions for the Site
        /// </summary>
        public RoleDefinitionCollection RoleDefinitions
        {
            get { return this._roleDefinitions; }
            private set { this._roleDefinitions = value; }
        }

        /// <summary>
        /// List of Role Assignments for the Site
        /// </summary>
        public RoleAssignmentCollection RoleAssignments
        {
            get { return this._roleAssignments; }
            private set { this._roleAssignments = value; }
        }

        #endregion
    }
}
