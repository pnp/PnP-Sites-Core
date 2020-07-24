using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.AzureActiveDirectory
{
    /// <summary>
    /// Collection of AAD Users' Licenses
    /// </summary>
    public partial class UserLicenseCollection : BaseProvisioningTemplateObjectCollection<UserLicense>
    {
        /// <summary>
        /// Constructor for UserLicenseCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public UserLicenseCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
