using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Collection of Templates for Microsoft Teams
    /// </summary>
    public partial class TeamTemplateCollection : BaseProvisioningTemplateObjectCollection<TeamTemplate>
    {
        /// <summary>
        /// Constructor for TeamTemplateCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TeamTemplateCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
