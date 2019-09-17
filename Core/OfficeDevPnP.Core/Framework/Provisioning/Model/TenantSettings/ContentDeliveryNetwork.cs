using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the CDN settings to provision
    /// </summary>
    public partial class ContentDeliveryNetwork : BaseModel, IEquatable<ContentDeliveryNetwork>
    {
        #region Private Members

        private CdnSettings _publicCdn;
        private CdnSettings _privateCdn;

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public ContentDeliveryNetwork()
        {
        }

        /// <summary>
        /// Custom constructor with both public and private CDN settings
        /// </summary>
        public ContentDeliveryNetwork(CdnSettings publicCdn, CdnSettings privateCdn)
        {
            this._publicCdn = publicCdn;
            this._privateCdn = privateCdn;
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the Public CDN settings to provision
        /// </summary>
        public CdnSettings PublicCdn
        {
            get { return this._publicCdn; }
            private set
            {
                if (this._publicCdn != null)
                {
                    this._publicCdn.ParentTemplate = null;
                }
                this._publicCdn = value;
                if (this._publicCdn != null)
                {
                    this._publicCdn.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        /// <summary>
        /// Defines the Private CDN settings to provision
        /// </summary>
        public CdnSettings PrivateCdn
        {
            get { return this._privateCdn; }
            private set
            {
                if (this._privateCdn != null)
                {
                    this._privateCdn.ParentTemplate = null;
                }
                this._privateCdn = value;
                if (this._privateCdn != null)
                {
                    this._privateCdn.ParentTemplate = this.ParentTemplate;
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
                this.PublicCdn?.GetHashCode() ?? 0,
                this.PrivateCdn?.GetHashCode() ?? 0
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with CDN
        /// </summary>
        /// <param name="obj">Object that represents CDN</param>
        /// <returns>true if the current object is equal to the CDN</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ContentDeliveryNetwork))
            {
                return (false);
            }
            return (Equals((ContentDeliveryNetwork)obj));
        }

        /// <summary>
        /// Compares Cdn object based on PublicCdn and PrivateCdn properties.
        /// </summary>
        /// <param name="other">Cdn object</param>
        /// <returns>true if the Cdn object is equal to the current object; otherwise, false.</returns>
        public bool Equals(ContentDeliveryNetwork other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.PublicCdn == other.PublicCdn &&
                this.PrivateCdn == other.PrivateCdn
                );
        }

        #endregion
    }
}
