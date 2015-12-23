using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class ListInstanceCollection : ProvisioningTemplateCollection<ListInstance>
    {
        public ListInstanceCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
