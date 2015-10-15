using OfficeDevPnP.Core.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// This class holds deprecated navigation related methods
    /// </summary>
    public static partial class NavigationExtensions
    {
        #region TO BE DEPRECATED IN DECEMBER 2015 RELEASE
        /// <summary>
        /// Deletes all Quick Launch nodes
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        [Obsolete("Use DeleteAllNavigationNodes(web, NavigationType.QuickLaunch)")]
        public static void DeleteAllQuickLaunchNodes(this Web web)
        {
            DeleteAllNavigationNodes(web, NavigationType.QuickLaunch);
        }

        #endregion
    }
}
