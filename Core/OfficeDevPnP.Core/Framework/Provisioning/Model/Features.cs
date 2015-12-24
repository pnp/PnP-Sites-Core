using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that is used in the Site Template for OOB Features
    /// </summary>
    public partial class Features: BaseModel
    {
        private FeatureCollection _siteFeatures;
        private FeatureCollection _webFeatures;

        #region Constructors
        public Features()
        {
            this._siteFeatures = new FeatureCollection(this.ParentTemplate);
            this._webFeatures = new FeatureCollection(this.ParentTemplate);
        }
        #endregion

        #region Properties
        /// <summary>
        /// A Collection of Features at the Site level
        /// </summary>
        public FeatureCollection SiteFeatures
        {
            get{ return this._siteFeatures; }
            private set { this._siteFeatures = value; }
        }

        /// <summary>
        /// A Collection of Features at the Web level
        /// </summary>
        public FeatureCollection WebFeatures
        {
            get { return this._webFeatures; }
            private set { this._webFeatures = value; }
        }

        #endregion
    }
}