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
    public class ProvisioningTenant : BaseModel, IEquatable<ProvisioningTenant>
    {
        #region Private Members

        private AppCatalog _appCatalog;
        private ContentDeliveryNetwork _cdn;
        private SiteDesignCollection _siteDesigns;
        private SiteScriptCollection _siteScripts;
        private StorageEntityCollection _storageEntities;
        private WebApiPermissionCollection _webApiPermissions;

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
            this.AppCatalog = AppCatalog;
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

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|",
                this.AppCatalog?.GetHashCode() ?? 0,
                this.ContentDeliveryNetwork?.GetHashCode() ?? 0,
                this.SiteDesigns.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.SiteScripts.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.StorageEntities.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
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
        /// Compares ProvisioningTenant object based on AppCatalog and Cdn properties.
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
                this.StorageEntities.DeepEquals(other.StorageEntities)
                );
        }

        #endregion
    }
}
