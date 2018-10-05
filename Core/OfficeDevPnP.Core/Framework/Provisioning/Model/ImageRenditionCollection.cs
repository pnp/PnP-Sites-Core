using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of ImageRendition objects
    /// </summary>
    public partial class ImageRenditionCollection : BaseProvisioningTemplateObjectCollection<ImageRendition>
    {
        /// <summary>
        /// Constructor for ImageRenditionCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ImageRenditionCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
