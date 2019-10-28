using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.ContentTypes
{
    public class ApplyContentTypeConfiguration
    {
        [JsonProperty("provisionContentTypesToSubWebs")]
        public bool ProvisionContentTypesToSubWebs { get; set; }
    }
}
