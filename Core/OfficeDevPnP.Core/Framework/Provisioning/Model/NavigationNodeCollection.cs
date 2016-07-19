using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of NavigationNode objects
    /// </summary>
    public partial class NavigationNodeCollection : ProvisioningTemplateCollection<NavigationNode>
    {
        public NavigationNodeCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
