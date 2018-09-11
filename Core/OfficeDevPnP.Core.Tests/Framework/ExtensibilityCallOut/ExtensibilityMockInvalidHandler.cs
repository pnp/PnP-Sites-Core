using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Extensibility;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.ExtensibilityCallOut
{
    /// <summary>
    /// This mock simulates an invalid ExtensibilityHandler.
    /// There are at least two situations that will lead to invalid extensibility providers.
    /// 1. The class does not inherit from one of the required interfaces.
    /// 2. If the extensibility provider is built against a different version of the currently
    ///     executing OfficeDevPnP.Core assembly (for instance in a PowerShell session)
    /// </summary>
    public class ExtensibilityMockInvalidHandler
    {
    }
}
