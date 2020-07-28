using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    /// <summary>
    /// Defines a provisioning engine Token. Make sure to only use the TokenContext property to execute queries in token methods.
    /// </summary>
    public abstract class VolatileTokenDefinition : TokenDefinition
    {
        public VolatileTokenDefinition(Web web, params string[] token) : base(web, token)
        {
        }

        public void ClearVolatileCache(Web web)
        {
            this.CacheValue = null;
            this.Web = web;
        }
    }
}