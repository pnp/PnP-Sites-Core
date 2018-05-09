namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Types of client side pages headers that a page can use
    /// </summary>
    public enum ClientSidePageHeaderType
    {
        /// <summary>
        /// The page uses the default page header
        /// </summary>
        Default = 0,
        /// <summary>
        /// The page does not have a header
        /// </summary>
        None = 1,
        /// <summary>
        /// The page use a customized header (e.g. with image + offset)
        /// </summary>
        Custom = 2
    }
#endif
}
