using OfficeDevPnP.Core.Utilities.Context;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that holds the extension methods used to "tag" a client context for cloning support
    /// </summary>
    static partial class InternalClientContextExtensions
    {
        private const string PnPSettingsKey = "SharePointPnP$Settings$ContextCloning";

        internal static void AddContextSettings(this ClientRuntimeContext clientContext, ClientContextSettings contextData)
        {
            clientContext.StaticObjects[PnPSettingsKey] = contextData;
        }

        internal static ClientContextSettings GetContextSettings(this ClientRuntimeContext clientContext)
        {
            if (!clientContext.StaticObjects.TryGetValue(PnPSettingsKey, out object settingsObject))
            {
                return null;
            }

            return (ClientContextSettings)settingsObject;
        }
    }
}
