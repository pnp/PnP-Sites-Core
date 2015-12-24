using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Term objects
    /// </summary>
    public partial class TermCollection : ProvisioningTemplateCollection<Term>
    {
        public TermCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
