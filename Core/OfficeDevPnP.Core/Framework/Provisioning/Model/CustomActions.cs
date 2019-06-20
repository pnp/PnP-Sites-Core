using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object that represents a Collections of Custom Actions
    /// </summary>
    public partial class CustomActions : BaseModel
    {
        #region Private Members
        private CustomActionCollection _siteCustomActions;
        private CustomActionCollection _webCustomActions;
        #endregion

        #region Constructor
        /// <summary>
        /// Constructor for CustomActions class
        /// </summary>
        public CustomActions()
        {
            _siteCustomActions = new CustomActionCollection(this.ParentTemplate);
            _webCustomActions = new CustomActionCollection(this.ParentTemplate);
        }
        #endregion

        #region Properties
        /// <summary>
        /// A Collection of CustomActions at the Site level
        /// </summary>
        public CustomActionCollection SiteCustomActions
        {
            get { return this._siteCustomActions; }
            private set { this._siteCustomActions = value; }
        }

        /// <summary>
        /// A Collection of CustomActions at the Web level
        /// </summary>
        public CustomActionCollection WebCustomActions
        {
            get { return this._webCustomActions; }
            private set { this._webCustomActions = value; }
        }

        #endregion
    }
}
