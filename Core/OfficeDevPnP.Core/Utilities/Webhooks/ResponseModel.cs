#if !ONPREMISES
using Newtonsoft.Json;
using System.Collections.Generic;

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
#endif