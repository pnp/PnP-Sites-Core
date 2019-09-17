using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.SharePoint.InformationArchitecture
{
    /// <summary>
    /// Collection of DataRowAttachment objects
    /// </summary>
    public partial class DataRowAttachmentCollection : BaseProvisioningTemplateObjectCollection<DataRowAttachment>
    {
        /// <summary>
        /// Constructor for DataRowAttachmentCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public DataRowAttachmentCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
