using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Basic interface for any CanProvision Rule
    /// </summary>
    interface ICanProvisionRuleBase
    {
        /// <summary>
        /// The name of the CanProvision Rule
        /// </summary>
        String Name { get; }
    }
}
