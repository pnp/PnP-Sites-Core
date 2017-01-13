using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    public enum ProvisioningMessageType
    {
        Progress = 0,
        Error = 1,
        Warning = 2,
        Completed = 3,
        EasterEgg = 100
    }
}
