using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Tenant
{
    public class ApplyTenantConfiguration
    {
        [JsonProperty("doNotWaitForSitesToBeFullyCreated")]
        public bool DoNotWaitForSitesToBeFullyCreated { get; set; }

        [JsonIgnore]
        [Obsolete("Use DoNotWaitForSitesToBeFullyCreated")]
        public int DelayAfterModernSiteCreation { get; set; }
    }
}
