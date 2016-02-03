using System;
using Microsoft.SharePoint.Client;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Domain Object for custom actions  associated with a SharePoint list, Web site, or subsite.
    /// </summary>
    public partial class CustomAction : BaseModel, IEquatable<CustomAction>
    {
        #region Private Members
        private int _rightsValue = 0;
        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the value that specifies the permissions needed for the custom action.
        /// <seealso>
        ///     <cref>https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.permissionkind.aspx</cref>
        /// </seealso>
        /// </summary>
        [Obsolete("Use Rights")]
        public int RightsValue
        {
            get
            {
                return this._rightsValue;
            }
            set
            {
                this._rightsValue = value;
                BasePermissions _bp = new BasePermissions();
                if (Enum.IsDefined(typeof(PermissionKind), value))
                {
                    var _pk = (PermissionKind)value;
                    _bp.Set(_pk);
                    this.Rights = _bp;
                }
            }
        }

        #endregion
    }
}
