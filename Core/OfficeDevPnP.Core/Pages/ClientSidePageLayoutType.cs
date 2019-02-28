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
        Article,
        /// <summary>
        /// Home page of modern team sites
        /// </summary>
        Home,
#if !SP2019
        /// <summary>
        /// Page is an app page, hosting a single SPFX web part full screen
        /// </summary>
        SingleWebPartAppPage,
#endif
        /// <summary>
        /// Page is a repost / link page
        /// </summary>
        RepostPage
    }
#endif
    }
