using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model
{
    /// <summary>
    /// Defines the mapping between original file names and OpenXML file names
    /// </summary>
    public class PnPFilesMap
    {
        /// <summary>
        /// Key and value containing mapping details
        /// </summary>
        public Dictionary<String, String> Map { get; set; }

        /// <summary>
        /// Constructor for PnPFilesMap class
        /// </summary>
        public PnPFilesMap()
        {
            this.Map = new Dictionary<String, String>();
        }

        /// <summary>
        /// Constructor for PnPFilesMap class
        /// </summary>
        /// <param name="items">Items</param>
        public PnPFilesMap(Dictionary<String, String> items)
        {
            this.Map = items;
        }
    }
}
