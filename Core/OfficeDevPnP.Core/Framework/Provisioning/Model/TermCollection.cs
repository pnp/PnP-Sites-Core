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
        /// <summary>
        /// Constructor for TermCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TermCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
