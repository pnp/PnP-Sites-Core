using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Webhooks
{

    /// <summary>
    /// 
    /// </summary>
    internal class ResponseModel<T>
    {

        [JsonProperty(PropertyName = "value")]
        public List<T> Value { get; set; }
    }
}
