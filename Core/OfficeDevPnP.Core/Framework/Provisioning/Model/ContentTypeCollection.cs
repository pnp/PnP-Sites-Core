using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of ContentType objects
    /// </summary>
    public partial class ContentTypeCollection : ProvisioningTemplateCollection<ContentType>
    {
        /// <summary>
        /// Constructor for ContentTypeCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ContentTypeCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
