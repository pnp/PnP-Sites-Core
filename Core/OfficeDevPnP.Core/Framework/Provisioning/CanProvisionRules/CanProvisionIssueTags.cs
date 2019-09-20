using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// List of Tags for CanProvision Issues
    /// </summary>
    public enum CanProvisionIssueTags
    {
        /// <summary>
        /// The App Catalog is missing
        /// </summary>
        MISSING_APP_CATALOG,
        /// <summary>
        /// The App Catalog is there, but the user doesn't have proper permissions
        /// </summary>
        MISSING_APP_CATALOG_PERMISSIONS,
        /// <summary>
        /// Lack of Permissions to access the TermStore
        /// </summary>
        MISSING_TERMSTORE_PERMISSIONS,
        /// <summary>
        /// Lack of Permissions, the user is not a Tenant Admin, which is required by the rule
        /// </summary>
        USER_IS_NOT_TENANT_ADMIN,
        /// <summary>
        /// The App Catalog needs a few hours to be fully provisioned
        /// </summary>
        APP_CATALOG_NOT_YEY_FULLY_PROVISIONED
    }
}
