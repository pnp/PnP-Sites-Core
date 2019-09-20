using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.AzureActiveDirectory
{
    /// <summary>
    /// Collection of AAD Users
    /// </summary>
    public partial class UserCollection : BaseProvisioningTemplateObjectCollection<User>
    {
        /// <summary>
        /// Constructor for UserCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public UserCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
