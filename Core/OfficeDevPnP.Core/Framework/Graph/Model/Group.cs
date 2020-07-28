using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Defines a Microsoft Graph Group
    /// </summary>
    public class Group
    {
        /// <summary>
        /// True if the group is not displayed in certain parts of the Outlook UI: the Address Book, address lists for selecting message recipients, and the Browse Groups dialog for searching groups; otherwise, false. Default value is false.
        /// </summary>
        [JsonProperty("hideFromAddressLists", NullValueHandling = NullValueHandling.Ignore)]
        public bool? HideFromAddressLists { get; set; }

        /// <summary>
        /// True if the group is not displayed in Outlook clients, such as Outlook for Windows and Outlook on the web; otherwise, false. Default value is false.
        /// </summary>
        [JsonProperty("hideFromOutlookClients", NullValueHandling = NullValueHandling.Ignore)]
        public bool? HideFromOutlookClients { get; set; }
    }
}
