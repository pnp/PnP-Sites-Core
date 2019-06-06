using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Holds properties of Structural Navigation
    /// </summary>
    public class StructuralNavigationEntity
    {
        /// <summary>
        /// Default Constructor
        /// </summary>
        public StructuralNavigationEntity()
        {
            MaxDynamicItems = 20;
            ShowSubsites = true;
            ShowPages = false;
        }
        /// <summary>
        /// Site navigation is inherited from parent web
        /// </summary>
        public bool InheritFromParentWeb { get; set; }
        /// <summary>
        /// Site navigation powered by the SharePoint managed metadata service (taxonomy). Use it to build site navigation derived from a managed metadata taxonomy. Managed navigation often works best with the product catalog
        /// </summary>
        public bool ManagedNavigation { get; set; }
        /// <summary>
        /// Subsites will be displayed in navigation
        /// </summary>
        public bool ShowSubsites { get; set; }
        /// <summary>
        /// Pages will be displayed in navigation
        /// </summary>
        public bool ShowPages { get; set; }
        /// <summary>
        /// To display maximum number of items in the navigation
        /// </summary>
        public int MaxDynamicItems { get; set; }
        /// <summary>
        /// Display the current site, the nav items below the current site, and the site's siblings
        /// </summary>
        public bool ShowSiblings { get; set; }

    }
}
