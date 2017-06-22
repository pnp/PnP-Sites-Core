using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Enums
{
    /// <summary>
    /// Enums defining type of sorting for Structural Navigation
    /// </summary>
    public enum StructuralNavigationSorting
    {
        /// <summary>
        /// Automatically sort
        /// </summary>
        Automatically = 0,
        /// <summary>
        /// Sort Pages automatically and rest manually
        /// </summary>
        ManuallyButPagesAutomatically = 1,
        /// <summary>
        /// Manuallt sort
        /// </summary>
        Manually = 2
    }
}
