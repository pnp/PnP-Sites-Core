namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Types of client side pages that can be created
    /// </summary>
    public enum ClientSidePageLayoutType
    {
        /// <summary>
        /// Custom article page, used for user created pages
        /// </summary>
        Article,
        /// <summary>
        /// Home page of modern team sites
        /// </summary>
        Home
    }
#endif
}
