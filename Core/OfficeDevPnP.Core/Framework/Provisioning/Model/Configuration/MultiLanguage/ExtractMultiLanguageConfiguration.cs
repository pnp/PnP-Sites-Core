using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.MultiLanguage
{
    public class ExtractMultiLanguageConfiguration
    {
        [JsonProperty("persistMultilanguageResources")]
        public bool PersistResources { get; set; }

        [JsonProperty("resourceFilePrefix")]
        public string ResourceFilePrefix { get; set; }
    }
}
