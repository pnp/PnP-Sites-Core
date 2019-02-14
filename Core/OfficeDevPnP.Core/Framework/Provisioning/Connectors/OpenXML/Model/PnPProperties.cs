using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Connectors.OpenXML.Model
{
    /// <summary>
    /// Properties of the PnP OpenXML container
    /// </summary>
    [Serializable]
    public class PnPProperties
    {
        /// <summary>
        /// Unique ID for the PnP OpenXML file
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();

        /// <summary>
        /// Author of the PnP OpenXML file
        /// </summary>
        public String Author { get; set; }

        /// <summary>
        /// Date and Time of creation for the PnP OpenXML file
        /// </summary>
        public DateTime CreationDateTime { get; set; } = DateTime.Now;

        /// <summary>
        /// Name of the Generator (engine) of the PnP OpenXML file
        /// </summary>
        public String Generator { get; set; }

        public String TemplateFileName { get; set; }
    }
}
