using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class RoleDefinitionCollection : ProvisioningTemplateCollection<RoleDefinition>
    {
        public RoleDefinitionCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
