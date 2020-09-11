using Newtonsoft.Json;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Defines a Microsoft Graph user delta
    /// </summary>
    public class UserDelta
    {
        /// <summary>
        /// User objects with changes or all users if no SkipToken has been provided
        /// </summary>
        [JsonProperty("value", NullValueHandling = NullValueHandling.Ignore)]
        public IList<User> Users { get; set; }

        /// <summary>
        /// The DeltaToken which can be used when querying for changes to request changes made to User objects since this DeltaToken has been given out
        /// </summary>
        [JsonProperty("@odata.deltaLink", NullValueHandling = NullValueHandling.Ignore)]
        public string DeltaToken { get; set; }

        /// <summary>
        /// The NextLink which indicates there are more results
        /// </summary>
        [JsonProperty("@odata.nextLink", NullValueHandling = NullValueHandling.Ignore)]
        public string NextLink { get; set; }
    }
}
