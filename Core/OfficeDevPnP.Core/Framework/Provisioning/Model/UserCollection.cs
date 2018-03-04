using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of User objects
    /// </summary>
    public partial class UserCollection : ProvisioningTemplateCollection<User>
    {
        /// <summary>
        /// Constructor for UserCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public UserCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
