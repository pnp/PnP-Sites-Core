using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of TermLabel objects
    /// </summary>
    public partial class TermLabelCollection : ProvisioningTemplateCollection<TermLabel>
    {
        /// <summary>
        /// Constructor for TermLabelCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TermLabelCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
