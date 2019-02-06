using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    internal static class PnPUtility
    {
        /// <summary>
        /// Checks if the ClientObject is null
        /// </summary>        
        /// <param name="clientObject">Object to check</param>
        /// <returns>True if the server object is null, otherwise false</returns>
        internal static bool ServerObjectIsNull(ClientObject clientObject)
        {
            if (clientObject == null)
            {
                return true;
            }
            else if (!clientObject.ServerObjectIsNull.HasValue)
            {
                return false;
            }

            return clientObject.ServerObjectIsNull.Value;
        }
    }
}
