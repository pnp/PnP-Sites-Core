using System;
using System.Linq;
using System.Collections.Generic;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that is used in the site template
    /// </summary>
    public partial class SiteSecurity : BaseModel, IEquatable<SiteSecurity>
    {
        #region Private Members

        private UserCollection _additionalAdministrators;
        private UserCollection _additionalOwners;
        private UserCollection _additionalMembers;
        private UserCollection _additionalVisitors;
        private SiteGroupCollection _siteGroups;
        private SiteSecurityPermissions _permissions = new SiteSecurityPermissions();

        #endregion

        #region Constructor
        /// <summary>
        /// Constructor for SiteSecurity class
        /// </summary>
        public SiteSecurity()
        {
            this._additionalAdministrators = new UserCollection(this.ParentTemplate);
            this._additionalOwners = new UserCollection(this.ParentTemplate);
            this._additionalMembers = new UserCollection(this.ParentTemplate);
            this._additionalVisitors = new UserCollection(this.ParentTemplate);
            this._siteGroups = new SiteGroupCollection(this.ParentTemplate);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// A Collection of users that are associated as site collection adminsitrators
        /// </summary>
        public UserCollection AdditionalAdministrators
        {
            get { return _additionalAdministrators; }
            private set { _additionalAdministrators = value; }
        }

        /// <summary>
        /// Declares whether to clear existing administrators before adding new ones
        /// </summary>
        public Boolean ClearExistingAdministrators { get; set; }

        /// <summary>
        /// A Collection of users that are associated to the sites owners group
        /// </summary>
        public UserCollection AdditionalOwners
        {
            get { return _additionalOwners; }
            private set { _additionalOwners = value; }
        }

        /// <summary>
        /// Declares whether to clear existing owners before adding new ones
        /// </summary>
        public Boolean ClearExistingOwners { get; set; }

        /// <summary>
        /// A Collection of users that are associated to the sites members group
        /// </summary>
        public UserCollection AdditionalMembers
        {
            get { return _additionalMembers; }
            private set { _additionalMembers = value; }
        }

        /// <summary>
        /// Declares whether to clear existing members before adding new ones
        /// </summary>
        public Boolean ClearExistingMembers { get; set; }

        /// <summary>
        /// A Collection of users taht are associated to the sites visitors group
        /// </summary>
        public UserCollection AdditionalVisitors
        {
            get { return _additionalVisitors; }
            private set { _additionalVisitors = value; }
        }

        /// <summary>
        /// Declares whether to clear existing visitors before adding new ones
        /// </summary>
        public Boolean ClearExistingVisitors { get; set; }

        /// <summary>
        /// List of additional Groups for the Site
        /// </summary>
        public SiteGroupCollection SiteGroups
        {
            get { return _siteGroups; }
            private set { _siteGroups = value; }
        }

        /// <summary>
        /// List of Site Security Permissions for the Site
        /// </summary>
        public SiteSecurityPermissions SiteSecurityPermissions
        {
            get { return _permissions; }
            private set
            {
                if (this._permissions != null)
                {
                    this._permissions.ParentTemplate = null;
                }
                this._permissions = value;
                if (this._permissions != null)
                {
                    this._permissions.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Declares whether the to break role inheritance for the site, if it is a sub-site
        /// </summary>
        public Boolean BreakRoleInheritance { get; set; } = false;

        /// <summary>
        /// Declares whether to reset the role inheritance or not for the site, if it is a sub-site
        /// </summary>
        public Boolean ResetRoleInheritance { get; set; } = false;

        /// <summary>
        /// Defines whether to copy role assignments or not while breaking role inheritance
        /// </summary>
        public Boolean CopyRoleAssignments { get; set; } = false;

        /// <summary>
        /// Defines whether to remove unique role assignments or not if the site already breaks role inheritance. If true all existing unique role assignments on the site will be removed if BreakRoleInheritance also is true.
        /// </summary>
        public Boolean RemoveExistingUniqueRoleAssignments { get; set; } = false;

        /// <summary>
        /// Defines whether to clear subscopes or not while breaking role inheritance for the site
        /// </summary>
        public Boolean ClearSubscopes { get; set; } = false;

        /// <summary>
        /// Specifies the list of groups that are associated with the Web site. Groups in this list will appear under the Groups section in the People and Groups page.
        /// </summary>
        public String AssociatedGroups { get; set; }

        /// <summary>
        /// Specifies the default owners group for this site. The group will automatically be added to the end of the Associated Groups list.
        /// </summary>
        public String AssociatedOwnerGroup { get; set; }

        /// <summary>
        /// Specifies the default members group for this site. The group will automatically be added to the top of the Associated Groups list.
        /// </summary>
        public String AssociatedMemberGroup { get; set; }

        /// <summary>
        /// Specifies the default visitors group for this site. The group will automatically be added to the end of the Associated Groups list.
        /// </summary>
        public String AssociatedVisitorGroup { get; set; }
        
        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}",
                this.AdditionalAdministrators.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.AdditionalOwners.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.AdditionalMembers.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.AdditionalVisitors.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.SiteGroups.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.SiteSecurityPermissions != null ? this.SiteSecurityPermissions.RoleAssignments.Aggregate(0, (acc, next) => acc += next.GetHashCode()) : 0),
                (this.SiteSecurityPermissions != null ? this.SiteSecurityPermissions.RoleDefinitions.Aggregate(0, (acc, next) => acc += next.GetHashCode()) : 0),
                this.BreakRoleInheritance.GetHashCode(),
                this.CopyRoleAssignments.GetHashCode(),
                this.ClearSubscopes.GetHashCode(),
                this.ResetRoleInheritance.GetHashCode(),
                this.RemoveExistingUniqueRoleAssignments.GetHashCode(),
                this.AssociatedGroups.GetHashCode(),
                this.AssociatedOwnerGroup.GetHashCode(),
                this.AssociatedMemberGroup.GetHashCode(),
                this.AssociatedVisitorGroup.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with SiteSecurity
        /// </summary>
        /// <param name="obj">Object that represents SiteSecurity</param>
        /// <returns>true if the current object is equal to the SiteSecurity</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is SiteSecurity))
            {
                return (false);
            }
            return (Equals((SiteSecurity)obj));
        }

        /// <summary>
        /// Compares SiteSecurity object based on AdditionalAdministrators, AdditionalOwners, AdditionalMembers, AdditionalVisitors, SiteGroups, 
        /// SiteSecurityPermissions, BreakRoleInheritance, CopyRoleAssignments and ClearSubscopes properties.
        /// </summary>
        /// <param name="other">SiteSecurity object</param>
        /// <returns>true if the SiteSecurity object is equal to the current object; otherwise, false.</returns>
        public bool Equals(SiteSecurity other)
        {
            if (other == null)
            {
                return (false);
            }

            return (
                this.AdditionalAdministrators.DeepEquals(other.AdditionalAdministrators) &&
                this.AdditionalOwners.DeepEquals(other.AdditionalOwners) &&
                this.AdditionalMembers.DeepEquals(other.AdditionalMembers) &&
                this.AdditionalVisitors.DeepEquals(other.AdditionalVisitors) &&
                this.SiteGroups.DeepEquals(other.SiteGroups) &&
                (this.SiteSecurityPermissions != null ? this.SiteSecurityPermissions.RoleAssignments.DeepEquals(other.SiteSecurityPermissions.RoleAssignments) : true) &&
                (this.SiteSecurityPermissions != null ? this.SiteSecurityPermissions.RoleDefinitions.DeepEquals(other.SiteSecurityPermissions.RoleDefinitions) : true) &&
                this.BreakRoleInheritance == other.BreakRoleInheritance &&
                this.CopyRoleAssignments == other.CopyRoleAssignments &&
                this.ClearSubscopes == other.ClearSubscopes &&
                this.ResetRoleInheritance == other.ResetRoleInheritance &&
                this.RemoveExistingUniqueRoleAssignments == other.RemoveExistingUniqueRoleAssignments &&
                this.AssociatedGroups == other.AssociatedGroups &&
                this.AssociatedOwnerGroup == other.AssociatedOwnerGroup &&
                this.AssociatedMemberGroup == other.AssociatedMemberGroup &&
                this.AssociatedVisitorGroup == other.AssociatedVisitorGroup
                );
        }

        #endregion
    }
}
