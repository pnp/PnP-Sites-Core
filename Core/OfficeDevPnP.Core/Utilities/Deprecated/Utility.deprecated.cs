using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities
{
    public static partial class Utility
    {

        // TO BE DEPRECATED IN 2015 - 12 release
        /// <summary>
        /// Check if the property is loaded on the site object, if not the site object will be reloaded
        /// </summary>
        /// <param name="cc">Context to execute upon</param>
        /// <param name="site">Site to execute upon</param>
        /// <param name="propertyToCheck">Property to check</param>
        /// <returns>A reloaded site object</returns>
        [Obsolete("Use ClientObject.EnsureProperty() instead")]
        public static Site EnsureSite(ClientRuntimeContext cc, Site site, string propertyToCheck)
        {
            if (!site.IsObjectPropertyInstantiated(propertyToCheck))
            {
                // get instances to root web, since we are processing currently sub site 
                cc.Load(site);
                cc.ExecuteQueryRetry();
            }
            return site;
        }

        /// <summary>
        /// Check if the property is loaded on the web object, if not the web object will be reloaded
        /// </summary>
        /// <param name="cc">Context to execute upon</param>
        /// <param name="web">Web to execute upon</param>
        /// <param name="propertyToCheck">Property to check</param>
        /// <returns>A reloaded web object</returns>
        [Obsolete("Use ClientObject.EnsureProperty() instead")]
        public static Web EnsureWeb(ClientRuntimeContext cc, Web web, string propertyToCheck)
        {
            if (!web.IsObjectPropertyInstantiated(propertyToCheck))
            {
                // get instances to root web, since we are processing currently sub site 
                cc.Load(web);
                cc.ExecuteQueryRetry();
            }
            return web;
        }

    }
}
