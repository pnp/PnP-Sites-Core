using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{fqdn}",
     Description = "Returns a full url of the current host",
     Example = "{fqdn}",
     Returns = "mycompany.sharepoint.com")]
    public class FqdnToken: TokenDefinition
    {
        public FqdnToken(Web web): base(web, "{fqdn}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Web.EnsureProperty(w => w.Url);
                var uri = new Uri(TokenContext.Web.Url);
                CacheValue = $"{uri.DnsSafeHost.ToLower().Replace("-admin","")}";
            }
            return CacheValue;
        }
    }
}
