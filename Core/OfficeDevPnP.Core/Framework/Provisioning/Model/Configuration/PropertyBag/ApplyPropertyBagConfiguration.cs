using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.PropertyBag
{
    public class ApplyPropertyBagConfiguration
    {
        [JsonProperty("overwriteSystemValues")]
        public bool OverwriteSystemValues { get; set; }
    }
}
