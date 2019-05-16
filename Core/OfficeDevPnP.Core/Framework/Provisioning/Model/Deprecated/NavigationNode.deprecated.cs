using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Deprecated
{
    public partial class NavigationNode
    {
        /// <summary>
        /// Defines whether the Navigation Node for the Structural Navigation is visible or not
        /// </summary>
        [Obsolete("Removed because it is not used")]
        public Boolean IsVisible { get; set; }
    }
}
