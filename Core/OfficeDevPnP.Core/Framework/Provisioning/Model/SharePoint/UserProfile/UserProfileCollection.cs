using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.SPUPS
{
    /// <summary>
    /// Collection of UserProfile items
    /// </summary>
    public partial class UserProfileCollection : BaseProvisioningTemplateObjectCollection<UserProfile>
    {
        /// <summary>
        /// Constructor for UserProfileCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public UserProfileCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
