using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Extensibility
{
    public class ApplyExtensibilityConfiguration
    {
        [JsonProperty("handlers")]
        public List<ExtensibilityHandler> Handlers { get; set; } = new List<ExtensibilityHandler>();
    }
}
