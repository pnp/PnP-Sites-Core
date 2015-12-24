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
        public UserCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
