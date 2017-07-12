using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of ContentTypeBinding objects
    /// </summary>
    public partial class ContentTypeBindingCollection : ProvisioningTemplateCollection<ContentTypeBinding>
    {
        /// <summary>
        /// Constructor for ContentTypeBindingCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ContentTypeBindingCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
