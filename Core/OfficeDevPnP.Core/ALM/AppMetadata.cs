#if !ONPREMISES
using Newtonsoft.Json;
using System;

namespace OfficeDevPnP.Core.ALM
{
    /// <summary>
    /// App/solution metadata for apps stored in the corporate catalog
    /// </summary>
    public class AppMetadata
    {
        /// <summary>
        /// Unique ID of the library list item of the app/solution.
        /// </summary>
        [JsonProperty()]
        public Guid Id { get; internal set; }
        /// <summary>
        /// Returns version of the app / solution int the app catalog.
        /// </summary>
        [JsonProperty()]
        public Version AppCatalogVersion { get; internal set; }
        /// <summary>
        /// Returns whether an existing instance of an app/solution can be upgraded. 
        /// True if there's newer version available in app catalog compared to instance in site.
        /// </summary>
        [JsonProperty()]
        public bool CanUpgrade { get; internal set; }
        /// <summary>
        /// Returns whether app/solution has been deployed to the context site. 
        /// True if particular app/solution has been installed to the site.
        /// </summary>
        [JsonProperty()]
        public bool Deployed { get; internal set; }
        /// <summary>
        /// Returns version of the installed app/solution in the site context. 
        /// </summary>
        [JsonProperty()]
        public Version InstalledVersion { get; internal set; }
        /// <summary>
        /// Returns wheter app/solution is SharePoint Framework client-side solution. 
        /// True for SPFx, False for app/add-in.
        /// </summary>
        [JsonProperty()]
        public bool IsClientSideSolution { get; internal set; }
        /// <summary>
        /// Title of the solution
        /// </summary>
        [JsonProperty()]
        public string Title { get; internal set; }
    }
}
#endif