using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    public partial class TeamSecurityUserCollection : BaseProvisioningTemplateObjectCollection<TeamSecurityUser>
    {
        /// <summary>
        /// Constructor for TeamSecurityUserCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamSecurityUserCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
