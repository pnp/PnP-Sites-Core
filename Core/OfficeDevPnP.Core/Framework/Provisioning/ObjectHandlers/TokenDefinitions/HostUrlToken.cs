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
     Token = "{hosturl}",
     Description = "Returns a full url of the current host",
     Example = "{hosturl}",
     Returns = "https://mycompany.sharepoint.com")]
    public class HostUrlToken: TokenDefinition
    {
        public HostUrlToken(Web web): base(web, "{hosturl}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                this.Web.EnsureProperty(w => w.Url);
                var uri = new Uri(this.Web.Url);
                CacheValue = $"{uri.Scheme}://{uri.DnsSafeHost}";
            }
            return CacheValue;
        }
    }
}
