using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
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
