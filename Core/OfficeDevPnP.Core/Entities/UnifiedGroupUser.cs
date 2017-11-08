using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Defines a Unified Group user
    /// </summary>
    public class UnifiedGroupUser
    {
        /// <summary>
        /// Unified group user's user principal name
        /// </summary>
        public String UserPrincipalName { get; set; }
        /// <summary>
        /// Unified group user's display name
        /// </summary>
        public String DisplayName { get; set; }
        
    }
}
