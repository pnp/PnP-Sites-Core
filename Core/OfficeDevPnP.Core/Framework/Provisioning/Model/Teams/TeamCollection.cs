using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Collection of Teams for Microsoft Teams
    /// </summary>
    public partial class TeamCollection : BaseProvisioningTemplateObjectCollection<Team>
    {
        /// <summary>
        /// Constructor for TeamCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
