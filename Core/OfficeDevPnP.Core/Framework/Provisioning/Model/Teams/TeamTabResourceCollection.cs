using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Collection of Resources for Tabs in a Team Channel
    /// </summary>
    public partial class TeamTabResourceCollection : BaseProvisioningTemplateObjectCollection<TeamTabResource>
    {
        /// <summary>
        /// Constructor for TeamTabResourceCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamTabResourceCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
