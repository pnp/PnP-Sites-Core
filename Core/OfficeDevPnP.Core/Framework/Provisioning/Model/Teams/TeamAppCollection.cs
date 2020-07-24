using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Collection of Apps for Microsoft Teams
    /// </summary>
    public partial class TeamAppCollection : BaseProvisioningTemplateObjectCollection<TeamApp>
    {
        /// <summary>
        /// Constructor for TeamAppCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamAppCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
