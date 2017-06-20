using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of TermSete objects
    /// </summary>
    public partial class TermSetCollection : ProvisioningTemplateCollection<TermSet>
    {
        /// <summary>
        /// Constructor for TermSetCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TermSetCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
