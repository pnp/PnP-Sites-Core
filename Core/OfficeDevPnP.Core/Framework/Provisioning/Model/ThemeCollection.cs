using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of Theme objects
    /// </summary>
    public partial class ThemeCollection : BaseProvisioningTemplateObjectCollection<Theme>
    {
        /// <summary>
        /// Constructor for ThemeCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ThemeCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
