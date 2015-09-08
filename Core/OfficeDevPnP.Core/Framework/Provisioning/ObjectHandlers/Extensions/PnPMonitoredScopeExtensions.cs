using OfficeDevPnP.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{
    internal static class PnPMonitoredScopeExtensions
    {
        public static void LogPropertyUpdate(this PnPMonitoredScope scope, string propertyName)
        {
            scope.LogDebug(CoreResources.PnPMonitoredScopeExtensions_LogPropertyUpdate_Updating_property__0_, propertyName);
        }
    }
}
