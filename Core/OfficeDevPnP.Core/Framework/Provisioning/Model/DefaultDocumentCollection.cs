using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class DefaultDocumentCollection : ProvisioningTemplateCollection<DefaultDocument>
    {
        /// <summary>
        /// Constructor for DefaultDocumentCollection class.
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public DefaultDocumentCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
