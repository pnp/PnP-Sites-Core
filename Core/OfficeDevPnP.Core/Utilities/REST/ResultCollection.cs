using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.REST
{
    public class ResultCollection<T>
    {
        [JsonPropertyName("@odata.nextLink")]
        public string NextLink { get; set; }

        [JsonPropertyName("value")]
        public IEnumerable<T> Items { get; set; }
    }
}
