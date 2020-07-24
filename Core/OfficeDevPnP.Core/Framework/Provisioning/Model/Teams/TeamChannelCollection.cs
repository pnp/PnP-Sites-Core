using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines a collection of Channels for the Team
    /// </summary>
    public partial class TeamChannelCollection : BaseProvisioningTemplateObjectCollection<TeamChannel>
    {
        /// <summary>
        /// Constructor for TeamChannelCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamChannelCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
