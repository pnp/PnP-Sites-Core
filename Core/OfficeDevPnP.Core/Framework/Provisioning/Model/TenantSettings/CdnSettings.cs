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
    public partial class CdnSettings : BaseModel, IEquatable<CdnSettings>
    {
        #region Private Members

        private CdnOriginCollection _origins;

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

        /// <summary>
        /// Defines whether the CDN has to be enabled or disabled
        /// </summary>
        public Boolean Enabled { get; set; }

        /// <summary>
        /// Defines whether the CDN should have default origins
        /// </summary>
        public Boolean NoDefaultOrigins { get; set; }

        /// <summary>
        /// Defines the file extensions to include in the CDN policy.
        /// </summary>
        public String IncludeFileExtensions { get; set; }

        /// <summary>
        /// Defines the site classifications to exclude of the wild card origins.
        /// </summary>
        public String ExcludeRestrictedSiteClassifications { get; set; }

        /// <summary>
        /// Allows to opt-out of sites that have disabled NoScript.
        /// </summary>
        public String ExcludeIfNoScriptDisabled { get; set; }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                this.Origins.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.IncludeFileExtensions?.GetHashCode() ?? 0,
                this.ExcludeRestrictedSiteClassifications?.GetHashCode() ?? 0,
                this.ExcludeIfNoScriptDisabled?.GetHashCode() ?? 0
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
        /// Compares CdnSettings object based on Origins, IncludeFileExtensions, 
        /// ExcludeRestrictedSiteClassifications, ExcludeIfNoScriptDisabled properties.
        /// </summary>
        /// <param name="other">CdnSettings object</param>
        /// <returns>true if the CdnSettings object is equal to the current object; otherwise, false.</returns>
        public bool Equals(CdnSettings other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Origins.DeepEquals(other.Origins) &&
                this.IncludeFileExtensions == other.IncludeFileExtensions &&
                this.ExcludeRestrictedSiteClassifications == other.ExcludeRestrictedSiteClassifications &&
                this.ExcludeIfNoScriptDisabled == other.ExcludeIfNoScriptDisabled
                );
        }

        #endregion
    }
}
