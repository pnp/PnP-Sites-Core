using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Types of provisioning messages
    /// </summary>
    public enum ProvisioningMessageType
    {
        /// <summary>
        /// Value 0, represents provisioning is in progress
        /// </summary>
        Progress = 0,
        /// <summary>
        /// Value 1, represents provisioning generated error
        /// </summary>
        Error = 1,
        /// <summary>
        /// Value 2, represents provisioning generated warning
        /// </summary>
        Warning = 2,
        /// <summary>
        /// Value 3, represents provisioning is completed
        /// </summary>
        Completed = 3,
        /// <summary>
        /// Value 100, represents provisioning unexpected behaviour
        /// </summary>
        EasterEgg = 100
    }
}
