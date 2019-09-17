using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Collection of Tabs for a Channel in a Team
    /// </summary>
    public partial class TeamTabCollection : BaseProvisioningTemplateObjectCollection<TeamTab>
    {
        /// <summary>
        /// Constructor for TeamTabCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamTabCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
