using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that is used in the site template
    /// </summary>
    public partial class SiteSecurity : BaseModel
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
        /// A Collection of users that are associated to the sites owners group
        /// </summary>
        public UserCollection AdditionalOwners
        {
            get { return _additionalOwners; }
            private set { _additionalOwners = value; }
        }

        /// <summary>
        /// A Collection of users that are associated to the sites members group
        /// </summary>
        public UserCollection AdditionalMembers
        {
            get { return _additionalMembers; }
            private set { _additionalMembers = value; }
        }

        /// <summary>
        /// A Collection of users taht are associated to the sites visitors group
        /// </summary>
        public UserCollection AdditionalVisitors
        {
            get { return _additionalVisitors; }
            private set { _additionalVisitors = value; }
        }

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

        #endregion
    }
}
