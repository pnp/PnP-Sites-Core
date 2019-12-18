using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration.Lists
{
    public class ApplyListsConfiguration
    {
        [JsonProperty("ignoreDuplicateDataRowErrors")]
        public bool IgnoreDuplicateDataRowErrors { get; set; }
    }
}
