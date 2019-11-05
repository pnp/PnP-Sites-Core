namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Types of client side pages that can be created
    /// </summary>
    public enum ClientSidePageLayoutType
    {
        /// <summary>
        /// Custom article page, used for user created pages
        /// </summary>
        Article = 0,
        /// <summary>
        /// Home page of modern team sites
        /// </summary>
        Home = 1,
#if !SP2019
        /// <summary>
        /// Page is an app page, hosting a single SPFX web part full screen
        /// </summary>
        SingleWebPartAppPage = 2,
        /// <summary>
        /// Page is a custom search result page
        /// </summary>
        HeaderlessSearchResults = 4,
#endif
        /// <summary>
        /// Page is a repost / link page
        /// </summary>
        RepostPage = 3,
    }
#endif
}
