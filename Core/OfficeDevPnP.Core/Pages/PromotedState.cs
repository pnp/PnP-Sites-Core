namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Page promotion state
    /// </summary>
    public enum PromotedState
    {
        /// <summary>
        /// Regular client side page
        /// </summary>
        NotPromoted = 0,
        /// <summary>
        /// Page that will be promoted as news article after publishing
        /// </summary>
        PromoteOnPublish = 1,
        /// <summary>
        /// Page that is promoted as news article
        /// </summary>
        Promoted = 2
    }
#endif
}
