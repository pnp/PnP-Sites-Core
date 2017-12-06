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

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                this.AppCatalog?.GetHashCode() ?? 0,
                this.ContentDeliveryNetwork?.GetHashCode() ?? 0
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
                this.ContentDeliveryNetwork == other.ContentDeliveryNetwork
                );
        }

        #endregion
    }
}
