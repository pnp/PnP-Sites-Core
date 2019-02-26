using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Control data for controls of type 3 (= client side web parts) which persist using the data-sp-controldata property only
    /// </summary>
    public class ClientSideWebPartControlDataOnly : ClientSideWebPartControlData
    {
        [JsonProperty(PropertyName = "webPartData", NullValueHandling = NullValueHandling.Ignore)]
        public string WebPartData { get; set; }
    }
#endif
}
