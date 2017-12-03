using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines the CDN Settings for a CDN to provision
    /// </summary>
    public class CdnSettings : BaseModel, IEquatable<CdnSettings>
    {
        #region Private Members

        private CdnOriginCollection _origins;
        private String _includeFileExtensions;
        private String _excludeRestrictedSiteClassifications;
        private String _excludeIfNoScriptDisabled;

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public CdnSettings()
        {
            this._origins = new Model.CdnOriginCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Custom constructor
        /// </summary>
        public CdnSettings(CdnOriginCollection origins): base()
        {
            this._origins.AddRange(origins);
        }

        #endregion

        #region Public Members

        /// <summary>
        /// Defines the CDN Origins settings to provision
        /// </summary>
        public CdnOriginCollection Origins
        {
            get { return this._origins; }
            private set { this._origins = value; }
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|",
                this.Origins.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with CdnSettings
        /// </summary>
        /// <param name="obj">Object that represents CdnSettings</param>
        /// <returns>true if the current object is equal to the CdnSettings</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is CdnSettings))
            {
                return (false);
            }
            return (Equals((CdnSettings)obj));
        }

        /// <summary>
        /// Compares CdnSettings object based on Origins properties.
        /// </summary>
        /// <param name="other">CdnSettings object</param>
        /// <returns>true if the CdnSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(CdnSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Origins.DeepEquals(other.Origins)
                );
        }

        #endregion
    }
}
