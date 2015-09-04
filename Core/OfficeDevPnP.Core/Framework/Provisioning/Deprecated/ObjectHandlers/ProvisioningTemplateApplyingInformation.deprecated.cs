using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    public partial class ProvisioningTemplateApplyingInformation
    {
        [Obsolete("Please don't use this member, insted use MessagesDelegate. This method will be removed in the November release.")]
        public ProvisioningMessagesDelegate MessageDelegate
        {
            get { return (this.MessagesDelegate); }
            set { this.MessagesDelegate = value; }
        }
    }
}
