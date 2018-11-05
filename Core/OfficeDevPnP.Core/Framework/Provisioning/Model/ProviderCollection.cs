using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Provider objects
    /// </summary>
    public partial class ProviderCollection : BaseProvisioningTemplateObjectCollection<Provider>
    {
        /// <summary>
        /// Constructor for ProviderCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ProviderCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {
            
        }
    }
}
