using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines the Apps to install or update on the Team
    /// </summary>
    public partial class TeamAppInstanceCollection : BaseProvisioningTemplateObjectCollection<TeamAppInstance>
    {
        /// <summary>
        /// Constructor for TeamAppInstanceCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamAppInstanceCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
