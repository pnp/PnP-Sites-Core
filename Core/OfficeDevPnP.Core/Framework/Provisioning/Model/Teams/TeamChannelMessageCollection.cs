using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Collection of Messages for a Team Channel
    /// </summary>
    public partial class TeamChannelMessageCollection : BaseProvisioningTemplateObjectCollection<TeamChannelMessage>
    {
        /// <summary>
        /// Constructor for TeamChannelMessageCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamChannelMessageCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
