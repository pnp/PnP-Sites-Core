namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Client side webpart object (retrieved via the _api/web/GetClientSideWebParts REST call)
    /// </summary>
    public class ClientSideComponent
    {
        /// <summary>
        /// Component type for client side webpart object
        /// </summary>
        public int ComponentType { get; set; }
        /// <summary>
        /// Id for client side webpart object
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// Manifest for client side webpart object
        /// </summary>
        public string Manifest { get; set; }
        /// <summary>
        /// Manifest type for client side webpart object
        /// </summary>
        public int ManifestType { get; set; }
        /// <summary>
        /// Name for client side webpart object
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Status for client side webpart object
        /// </summary>
        public int Status { get; set; }
    }
#endif
}
