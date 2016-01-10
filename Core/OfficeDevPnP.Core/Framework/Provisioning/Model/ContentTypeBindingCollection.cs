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
        public ContentTypeBindingCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
