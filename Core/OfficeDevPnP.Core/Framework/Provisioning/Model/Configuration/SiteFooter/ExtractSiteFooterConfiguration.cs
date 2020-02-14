using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.SiteFooter
{
    public class ExtractSiteFooterConfiguration
    {
        [JsonProperty("removeExistingNodes")]
        public bool RemoveExistingNodes { get; set; }
    }
}
