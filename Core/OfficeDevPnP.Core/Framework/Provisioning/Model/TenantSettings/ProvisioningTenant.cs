using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the Tenant-wide settings to provision
    /// </summary>
    public partial class ProvisioningTenant : BaseModel, IEquatable<ProvisioningTenant>
    {
        #region Private Members

        private AppCatalog _appCatalog;
        private ContentDeliveryNetwork _cdn;
        private SiteDesignCollection _siteDesigns;
        private SiteScriptCollection _siteScripts;
        private StorageEntityCollection _storageEntities;
        private WebApiPermissionCollection _webApiPermissions;
        private ThemeCollection _themes;
        private Office365Groups.Office365GroupLifecyclePolicyCollection _office365GroupLifecyclePolicies;
        private SPUPS.UserProfileCollection _SPUsersProfiles;

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public ProvisioningTenant()
        {
        }

        /// <summary>
        /// Custom constructor which accepts AppCatalog and CDN settings
        /// </summary>
        public ProvisioningTenant(AppCatalog appCatalog, ContentDeliveryNetwork cdn)
        {
            this.AppCatalog = appCatalog;
            this.ContentDeliveryNetwork = cdn;
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the AppCatalog settings to provision
        /// </summary>
        public AppCatalog AppCatalog
        {
            get { return this._appCatalog; }
            private set
            {
                if (this._appCatalog != null)
                {
                    this._appCatalog.ParentTemplate = null;
                }
                this._appCatalog = value;
                if (this._appCatalog != null)
                {
                    this._appCatalog.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines the CDN settings to provision
        /// </summary>
        public ContentDeliveryNetwork ContentDeliveryNetwork
        {
            get { return this._cdn; }
            private set
            {
                if (this._cdn != null)
                {
                    this._cdn.ParentTemplate = null;
                }
                this._cdn = value;
                if (this._cdn != null)
                {
                    this._cdn.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Gets or sets SiteDesigns for the tenant
        /// </summary>
        public SiteDesignCollection SiteDesigns
        {
            get
            {
                if (this._siteDesigns == null)
                {
                    this._siteDesigns = new SiteDesignCollection(this.ParentTemplate);
                }
                return this._siteDesigns;
            }
            set
            {
                if (this._siteDesigns != null)
                {
                    this._siteDesigns.ParentTemplate = null;
                }
                this._siteDesigns = value;
                if (this._siteDesigns != null)
                {
                    this._siteDesigns.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Gets or sets SiteScripts for the tenant
        /// </summary>
        public SiteScriptCollection SiteScripts
        {
            get
            {
                if (this._siteScripts == null)
                {
                    this._siteScripts = new SiteScriptCollection(this.ParentTemplate);
                }
                return this._siteScripts;
            }
            set
            {
                if (this._siteScripts != null)
                {
                    this._siteScripts.ParentTemplate = null;
                }
                this._siteScripts = value;
                if (this._siteScripts != null)
                {
                    this._siteScripts.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Gets or sets StorageEntities for the tenant
        /// </summary>
        public StorageEntityCollection StorageEntities
        {
            get
            {
                if (this._storageEntities == null)
                {
                    this._storageEntities = new StorageEntityCollection(this.ParentTemplate);
                }
                return this._storageEntities;
            }
            set
            {
                if (this._storageEntities != null)
                {
                    this._storageEntities.ParentTemplate = null;
                }
                this._storageEntities = value;
                if (this._storageEntities != null)
                {
                    this._storageEntities.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Gets or sets StorageEntities for the tenant
        /// </summary>
        public WebApiPermissionCollection WebApiPermissions
        {
            get
            {
                if (this._webApiPermissions == null)
                {
                    this._webApiPermissions = new WebApiPermissionCollection(this.ParentTemplate);
                }
                return this._webApiPermissions;
            }
            set
            {
                if (this._webApiPermissions != null)
                {
                    this._webApiPermissions.ParentTemplate = null;
                }
                this._webApiPermissions = value;
                if (this._webApiPermissions != null)
                {
                    this._webApiPermissions.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Gets or sets StorageEntities for the tenant
        /// </summary>
        public ThemeCollection Themes
        {
            get
            {
                if (this._themes == null)
                {
                    this._themes = new ThemeCollection(this.ParentTemplate);
                }
                return this._themes;
            }
            set
            {
                if (this._themes != null)
                {
                    this._themes.ParentTemplate = null;
                }
                this._themes = value;
                if (this._themes != null)
                {
                    this._themes.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        public Office365Groups.Office365GroupsSettings Office365GroupsSettings { get; set; } = new Office365Groups.Office365GroupsSettings();

        /// <summary>
        /// Gets or sets Office365GroupLifecyclePolicies for the tenant
        /// </summary>
        public Office365Groups.Office365GroupLifecyclePolicyCollection Office365GroupLifecyclePolicies
        {
            get
            {
                if (this._office365GroupLifecyclePolicies == null)
                {
                    this._office365GroupLifecyclePolicies = new Office365Groups.Office365GroupLifecyclePolicyCollection(this.ParentTemplate);
                }
                return this._office365GroupLifecyclePolicies;
            }
            set
            {
                if (this._office365GroupLifecyclePolicies != null)
                {
                    this._office365GroupLifecyclePolicies.ParentTemplate = null;
                }
                this._office365GroupLifecyclePolicies = value;
                if (this._office365GroupLifecyclePolicies != null)
                {
                    this._office365GroupLifecyclePolicies.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Gets or sets SPUserProfiles for the tenant
        /// </summary>
        public SPUPS.UserProfileCollection SPUsersProfiles
        {
            get
            {
                if (this._SPUsersProfiles == null)
                {
                    this._SPUsersProfiles = new SPUPS.UserProfileCollection(this.ParentTemplate);
                }
                return this._SPUsersProfiles;
            }
            set
            {
                if (this._SPUsersProfiles != null)
                {
                    this._SPUsersProfiles.ParentTemplate = null;
                }
                this._SPUsersProfiles = value;
                if (this._SPUsersProfiles != null)
                {
                    this._SPUsersProfiles.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        public SharingSettings SharingSettings { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}",
                this.AppCatalog?.GetHashCode() ?? 0,
                this.ContentDeliveryNetwork?.GetHashCode() ?? 0,
                this.SiteDesigns.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.SiteScripts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.StorageEntities.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.WebApiPermissions.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Themes.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Office365GroupsSettings.GetHashCode(),
                this.Office365GroupLifecyclePolicies.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.SPUsersProfiles.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.SharingSettings.GetHashCode()
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with ProvisioningTenant
        /// </summary>
        /// <param name="obj">Object that represents ProvisioningTenant</param>
        /// <returns>true if the current object is equal to the ProvisioningTenant</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ProvisioningTenant))
            {
                return (false);
            }
            return (Equals((ProvisioningTenant)obj));
        }

        /// <summary>
        /// Compares ProvisioningTenant object based on AppCatalog, CDN, SiteDesigns, SiteScripts,
        /// StorageEntities, WebApiPermissions, Themes, Office365GroupsSettings, Office365GroupLifecyclePolicies,
        /// SPUserProfiles, and SharingSettings
        /// </summary>
        /// <param name="other">ProvisioningTenant object</param>
        /// <returns>true if the ProvisioningTenant object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ProvisioningTenant other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.AppCatalog == other.AppCatalog &&
                this.ContentDeliveryNetwork == other.ContentDeliveryNetwork &&
                this.SiteDesigns.DeepEquals(other.SiteDesigns) &&
                this.SiteScripts.DeepEquals(other.SiteScripts) &&
                this.StorageEntities.DeepEquals(other.StorageEntities) &&
                this.WebApiPermissions.DeepEquals(other.WebApiPermissions) &&
                this.Themes.DeepEquals(other.Themes) &&
                this.Office365GroupsSettings.Equals(other.Office365GroupsSettings) &&
                this.Office365GroupLifecyclePolicies.DeepEquals(other.Office365GroupLifecyclePolicies) &&
                this.SPUsersProfiles.DeepEquals(other.SPUsersProfiles) &&
                this.SharingSettings.Equals(other.SharingSettings)
                );
        }

        #endregion
    }
}
