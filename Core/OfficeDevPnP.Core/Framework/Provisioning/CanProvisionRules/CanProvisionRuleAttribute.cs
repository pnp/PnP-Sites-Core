using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Attribute to decorate a CanProvisionRule
    /// </summary>
    internal class CanProvisionRuleAttribute : Attribute
    {
        /// <summary>
        /// The Sequential position of the CanProvisionRule
        /// </summary>
        public Int32 Sequence { get; set; } = 0;

        /// <summary>
        /// The scope of the CanProvisionRule
        /// </summary>
        public CanProvisionScope Scope { get; set; } = CanProvisionScope.Site;
    }

    /// <summary>
    /// Enum with the list of scopes for CanProvisionRule instances
    /// </summary>
    internal enum CanProvisionScope
    {
        /// <summary>
        /// The scope is a single Site with a Provisioning Template
        /// </summary>
        Site,
        /// <summary>
        /// The scope is the whole SharePoint Tenant with a Provisioning Hierarchy
        /// </summary>
        Tenant,
        /// <summary>
        /// The scope is the whole Office 365 Tenant
        /// </summary>
        Office365,
    }
}
