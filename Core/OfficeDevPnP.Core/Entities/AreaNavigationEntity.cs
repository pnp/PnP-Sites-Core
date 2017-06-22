using OfficeDevPnP.Core.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Entity description navigation
    /// </summary>
    public class AreaNavigationEntity
    {

        /// <summary>
        /// Default Constructor
        /// </summary>
        public AreaNavigationEntity()
        {
            GlobalNavigation = new StructuralNavigationEntity();
            CurrentNavigation = new StructuralNavigationEntity();
        }

        /// <summary>
        /// Specifies the Global Navigation (top bar navigation)
        /// </summary>
        public StructuralNavigationEntity GlobalNavigation { get; set; }

        /// <summary>
        /// Specifies the Current Navigation (quick launch navigation)
        /// </summary>
        public StructuralNavigationEntity CurrentNavigation { get; set; }

        /// <summary>
        /// Defines the sorting
        /// </summary>
        public StructuralNavigationSorting Sorting { get; set; }

        /// <summary>
        /// Defines if sorted ascending
        /// </summary>
        public Boolean SortAscending { get; set; }

        /// <summary>
        /// Defines sorted by value
        /// </summary>
        public StructuralNavigationSortBy SortBy { get; set; }

        /// <summary>
        /// Defines if new pages should be added to navigation when using managed metadata navigation
        /// </summary>
        public Boolean AddNewPagesToNavigation { get; set; }

        /// <summary>
        /// Defines if friendly URLs should be created for new pages when using managed metadata navigation
        /// </summary>
        public Boolean CreateFriendlyUrlsForNewPages { get; set; }
    }
}
