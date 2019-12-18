namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// List of possible OOB web parts
    /// </summary>
    public enum DefaultClientSideWebParts
    {
        /// <summary>
        /// Third party webpart
        /// </summary>
        ThirdParty,
        /// <summary>
        /// Content Rollup webpart
        /// </summary>
        ContentRollup,
#if !ONPREMISES
        /// <summary>
        /// Bing Map webpart
        /// </summary>
        BingMap,
#endif
        /// <summary>
        /// Content Embed webpart
        /// </summary>
        ContentEmbed,
        /// <summary>
        /// Document Embed webpart
        /// </summary>
        DocumentEmbed,
        /// <summary>
        /// Image webpart
        /// </summary>
        Image,
        /// <summary>
        /// Image Gallery webpart
        /// </summary>
        ImageGallery,
        /// <summary>
        /// Link Preview webpart
        /// </summary>
        LinkPreview,
        /// <summary>
        /// News Feed webpart
        /// </summary>
        NewsFeed,
        /// <summary>
        /// News Reel webpart
        /// </summary>
        NewsReel,
#if !ONPREMISES
        /// <summary>
        /// News webpart (the "new" version of NewsReel) - they look the same but this one supports filtering properly
        /// </summary>
        News,
        /// <summary>
        /// PowerBI Report Embed webpart
        /// </summary>
        PowerBIReportEmbed,
#endif
        /// <summary>
        /// Quick Chart webpart
        /// </summary>
        QuickChart,
        /// <summary>
        /// Site Activity webpart
        /// </summary>
        SiteActivity,
        /// <summary>
        /// Video Embed webpart 
        /// </summary>
        VideoEmbed,
        /// <summary>
        /// Yammer Embed webpart
        /// </summary>
        YammerEmbed,
        /// <summary>
        /// Events webpart
        /// </summary>
        Events,
#if !ONPREMISES
        /// <summary>
        /// Group Calendar webpart
        /// </summary>
        GroupCalendar,
#endif
        /// <summary>
        /// Hero webpart
        /// </summary>
        Hero,
        /// <summary>
        /// List webpart
        /// </summary>
        List,
        /// <summary>
        /// Page Title webpart
        /// </summary>
        PageTitle,
        /// <summary>
        /// People webpart
        /// </summary>
        People,
        /// <summary>
        /// Quick Links webpart
        /// </summary>
        QuickLinks,
        /// <summary>
        /// Custom Message Region web part
        /// </summary>
        CustomMessageRegion,
        /// <summary>
        /// Divider web part
        /// </summary>
        Divider,
#if !ONPREMISES
        /// <summary>
        /// Microsoft Forms web part
        /// </summary>
        MicrosoftForms,
#endif
        /// <summary>
        /// Spacer web part
        /// </summary>
        Spacer,
#if !ONPREMISES
        /// <summary>
        /// Web part to host SharePoint Add-In parts
        /// </summary>
        ClientWebPart,
        /// <summary>
        /// Web part to host PowerApps
        /// </summary>
        PowerApps,
        /// <summary>
        /// Web part to show code
        /// </summary>
        CodeSnippet,
        /// <summary>
        /// Web part to show one or more properties of the page as page content
        /// </summary>
        PageFields,
        /// <summary>
        /// Weather web part
        /// </summary>
        Weather,
        /// <summary>
        /// YouTube embed web part
        /// </summary>
        YouTube,
        /// <summary>
        /// My documents web part
        /// </summary>
        MyDocuments,
        /// <summary>
        /// Yammer feed web part
        /// </summary>
        YammerFullFeed,
        /// <summary>
        /// CountDown web part
        /// </summary>
        CountDown,
        /// <summary>
        /// List properties web part
        /// </summary>
        ListProperties,
        /// <summary>
        /// MarkDown web part
        /// </summary>
        MarkDown,
        /// <summary>
        /// Planner web part
        /// </summary>
        Planner,
        /// <summary>
        /// Sites web part
        /// </summary>
        Sites,
        /// <summary>
        /// Call to Action web part
        /// </summary>
        CallToAction,
        /// <summary>
        /// Button web part
        /// </summary>
        Button
#endif
    }
#endif
    }
