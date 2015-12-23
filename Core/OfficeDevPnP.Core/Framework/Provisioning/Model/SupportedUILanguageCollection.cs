using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class SupportedUILanguages : ProvisioningTemplateCollection<SupportedUILanguage>
    {
        public SupportedUILanguages(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
