using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of RoleDefinition objects
    /// </summary>
    public partial class RoleDefinitionCollection : ProvisioningTemplateCollection<RoleDefinition>
    {
        public RoleDefinitionCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
