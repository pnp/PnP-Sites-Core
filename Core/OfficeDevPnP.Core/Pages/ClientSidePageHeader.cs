using AngleSharp.Parser.Html;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Diagnostics;
using System;
using System.Linq;
using System.Net;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Class that implements the client side page header
    /// </summary>
    public class ClientSidePageHeader
    {
#if SP2019
        private const string NoPageHeader      = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.3\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.3&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;0&#125;&#125;\"></div></div>";
        private const string DefaultPageHeader = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.3\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.3&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;4,&quot;translateX&quot;&#58;50,&quot;translateY&quot;&#58;50&#125;&#125;\"></div></div>";
        private const string CustomPageHeader  = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.3\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&quot;imageSource&quot;&#58;&quot;@@imageSource@@&quot;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.3&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;2,&quot;siteId&quot;&#58;&quot;@@siteId@@&quot;,&quot;webId&quot;&#58;&quot;@@webId@@&quot;,&quot;listId&quot;&#58;&quot;@@listId@@&quot;,&quot;uniqueId&quot;&#58;&quot;&#123;@@uniqueId@@&#125;&quot;@@focalPoints@@&#125;&#125;\"></div></div>";
#else
        private const string NoPageHeader = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.4\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;NoImage&quot;,&quot;textAlignment&quot;&#58;&quot;@@textalignment@@&quot;,&quot;showTopicHeader&quot;&#58;@@showtopicheader@@,&quot;showPublishDate&quot;&#58;@@showpublishdate@@,&quot;topicHeader&quot;&#58;&quot;@@topicheader@@&quot;&#125;&#125;\"></div></div>";
        private const string DefaultPageHeader = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.4\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;@@layouttype@@&quot;,&quot;textAlignment&quot;&#58;&quot;@@textalignment@@&quot;,&quot;showTopicHeader&quot;&#58;@@showtopicheader@@,&quot;showPublishDate&quot;&#58;@@showpublishdate@@,&quot;topicHeader&quot;&#58;&quot;@@topicheader@@&quot;,&quot;authorByline&quot;&#58;[@@authorbyline@@],&quot;authors&quot;&#58;[@@authors@@]&#125;&#125;\"></div></div>";
        private const string CustomPageHeader = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.4\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&quot;imageSource&quot;&#58;&quot;@@imageSource@@&quot;&#125;,&quot;links&quot;&#58;&#123;&#125;,&quot;customMetadata&quot;&#58;&#123;&quot;imageSource&quot;&#58;&#123;&quot;siteId&quot;&#58;&quot;@@siteId@@&quot;,&quot;webId&quot;&#58;&quot;@@webId@@&quot;,&quot;listId&quot;&#58;&quot;@@listId@@&quot;,&quot;uniqueId&quot;&#58;&quot;@@uniqueId@@&quot;&#125;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;2,&quot;layoutType&quot;&#58;&quot;@@layouttype@@&quot;,&quot;textAlignment&quot;&#58;&quot;@@textalignment@@&quot;,&quot;showTopicHeader&quot;&#58;@@showtopicheader@@,&quot;showPublishDate&quot;&#58;@@showpublishdate@@,&quot;topicHeader&quot;&#58;&quot;@@topicheader@@&quot;,&quot;authorByline&quot;&#58;[@@authorbyline@@],&quot;authors&quot;&#58;[@@authors@@],&quot;altText&quot;&#58;&quot;@@alternativetext@@&quot;,&quot;webId&quot;&#58;&quot;@@webId@@&quot;,&quot;siteId&quot;&#58;&quot;@@siteId@@&quot;,&quot;listId&quot;&#58;&quot;@@listId@@&quot;,&quot;uniqueId&quot;&#58;&quot;@@uniqueId@@&quot;@@focalPoints@@&#125;&#125;\"></div></div>";
#endif
        private ClientSidePageHeaderType pageHeaderType;
        private string imageServerRelativeUrl;
        private ClientContext clientContext;
        private bool headerImageResolved = false;
        private Guid siteId = Guid.Empty;
        private Guid webId = Guid.Empty;
        private Guid listId = Guid.Empty;
        private Guid uniqueId = Guid.Empty;

        /// <summary>
        /// Returns the type of header
        /// </summary>
        public ClientSidePageHeaderType Type
        {
            get
            {
                return this.pageHeaderType;
            }
        }

        /// <summary>
        /// Server relative link to page header image, set to null for default header image.
        /// Note: image needs to reside in the current site
        /// </summary>
        public string ImageServerRelativeUrl
        {
            get
            {
                return this.imageServerRelativeUrl;
            }
            set
            {
                this.imageServerRelativeUrl = value;
                this.headerImageResolved = false;
            }
        }

        /// <summary>
        /// Image focal point X coordinate
        /// </summary>
        public double? TranslateX { get; set; }

        /// <summary>
        /// Image focal point Y coordinate
        /// </summary>
        public double? TranslateY { get; set; }

        /// <summary>
        /// Type of layout used inside the header
        /// </summary>
        public ClientSidePageHeaderLayoutType LayoutType { get; set; }

#if !SP2019
        /// <summary>
        /// Alignment of the title in the header
        /// </summary>
        public ClientSidePageHeaderTitleAlignment TextAlignment { get; set; }

        /// <summary>
        /// Show the topic header in the title region
        /// </summary>
        public bool ShowTopicHeader { get; set; }

        /// <summary>
        /// Show the page publication date in the title region
        /// </summary>
        public bool ShowPublishDate { get; set; }

        /// <summary>
        /// The topic header text to show if ShowTopicHeader is set to true
        /// </summary>
        public string TopicHeader { get; set; }

        /// <summary>
        /// Alternative text for the header image
        /// </summary>
        public string AlternativeText { get; set; }

        /// <summary>
        /// Page author(s) to be displayed
        /// </summary>
        public string Authors { get; set; }

        /// <summary>
        /// Page author byline
        /// </summary>
        public string AuthorByLine { get; set; }

        /// <summary>
        /// Id of the page author
        /// </summary>
        public int AuthorByLineId { get; set; }
#endif

        #region construction
        /// <summary>
        /// Creates a custom header with a custom image
        /// </summary>
        /// <param name="cc">ClientContext of the site hosting the image</param>
        /// <param name="pageHeaderType">Type of page header</param>
        /// <param name="imageServerRelativeUrl">Server relative image url</param>
        public ClientSidePageHeader(ClientContext cc, ClientSidePageHeaderType pageHeaderType, string imageServerRelativeUrl)
        {
            this.imageServerRelativeUrl = imageServerRelativeUrl;
            this.clientContext = cc;
            this.pageHeaderType = pageHeaderType;
            this.LayoutType = ClientSidePageHeaderLayoutType.FullWidthImage;
#if !SP2019
            this.TextAlignment = ClientSidePageHeaderTitleAlignment.Left;
            this.ShowTopicHeader = false;
            this.TopicHeader = "";
            this.Authors = "";
            this.AlternativeText = "";
            this.ShowPublishDate = false;
            this.AuthorByLineId = -1;
#endif
        }

        /// <summary>
        /// Creates a custom header with a custom image + custom image offset
        /// </summary>
        /// <param name="cc">ClientContext of the site hosting the image</param>
        /// <param name="pageHeaderType">Type of page header</param>
        /// <param name="imageServerRelativeUrl">Server relative image url</param>
        /// <param name="translateX">X offset coordinate</param>
        /// <param name="translateY">Y offset coordinate</param>
        public ClientSidePageHeader(ClientContext cc, ClientSidePageHeaderType pageHeaderType, string imageServerRelativeUrl, double translateX, double translateY) : this(cc, pageHeaderType, imageServerRelativeUrl)
        {
            TranslateX = translateX;
            TranslateY = translateY;
        }
        #endregion

        /// <summary>
        /// Returns the header value to set a "no header"
        /// </summary>
        /// <param name="pageTitle">Title of the page</param>
        /// <param name="titleAlignment">Left align or center the title</param>
        /// <returns>Header html value that indicates "no header"</returns>
        private static string NoHeader(string pageTitle, ClientSidePageHeaderTitleAlignment titleAlignment)
        {
            if (pageTitle == null)
            {
                pageTitle = "";
            }
            else
            {
                pageTitle = EncodePageTitle(pageTitle);
            }

            string header = Replace1point4Defaults(NoPageHeader);

            return header.Replace("@@title@@", pageTitle).Replace("@@textalignment@@", titleAlignment.ToString());
        }

        /// <summary>
        /// Returns the header value to set a "no header"
        /// </summary>
        /// <param name="pageTitle">Title of the page</param>
        /// <returns>Header html value that indicates "no header"</returns>
        public static string NoHeader(string pageTitle)
        {
            return NoHeader(pageTitle, ClientSidePageHeaderTitleAlignment.Left);
        }

        /// <summary>
        /// Load the PageHeader object from the given html
        /// </summary>
        /// <param name="pageHeaderHtml">Page header html</param>
        public void FromHtml(string pageHeaderHtml)
        {
            // select all control div's
            if (String.IsNullOrEmpty(pageHeaderHtml))
            {
                this.pageHeaderType = ClientSidePageHeaderType.Default;
                return;
            }

            HtmlParser parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true });
            using (var document = parser.Parse(pageHeaderHtml))
            {
                var pageHeaderControl = document.All.Where(m => m.HasAttribute(CanvasControl.ControlDataAttribute)).FirstOrDefault();
                if (pageHeaderControl != null)
                {
                    var decoded = WebUtility.HtmlDecode(pageHeaderControl.GetAttribute(ClientSideWebPart.ControlDataAttribute));
                    JObject wpJObject = JObject.Parse(decoded);

                    // Store the server processed content as that's needed for full fidelity
                    if (wpJObject["serverProcessedContent"] != null)
                    {
                        if (wpJObject["serverProcessedContent"]["imageSources"] != null && wpJObject["serverProcessedContent"]["imageSources"]["imageSource"] != null)
                        {
                            this.imageServerRelativeUrl = wpJObject["serverProcessedContent"]["imageSources"]["imageSource"].ToString();
                        }

                        // Properties that apply to all header configurations
                        if (wpJObject["properties"]["layoutType"] != null)
                        {
                            this.LayoutType = (ClientSidePageHeaderLayoutType)Enum.Parse(typeof(ClientSidePageHeaderLayoutType), wpJObject["properties"]["layoutType"].ToString());
                        }
#if !SP2019
                        if (wpJObject["properties"]["textAlignment"] != null)
                        {
                            this.TextAlignment = (ClientSidePageHeaderTitleAlignment)Enum.Parse(typeof(ClientSidePageHeaderTitleAlignment), wpJObject["properties"]["textAlignment"].ToString());
                        }
                        if (wpJObject["properties"]["showTopicHeader"] != null)
                        {
                            bool showTopicHeader = false;
                            bool.TryParse(wpJObject["properties"]["showTopicHeader"].ToString(), out showTopicHeader);
                            this.ShowTopicHeader = showTopicHeader;
                        }
                        if (wpJObject["properties"]["showPublishDate"] != null)
                        {
                            bool showPublishDate = false;
                            bool.TryParse(wpJObject["properties"]["showPublishDate"].ToString(), out showPublishDate);
                            this.ShowPublishDate = showPublishDate;
                        }
                        if (wpJObject["properties"]["topicHeader"] != null)
                        {
                            this.TopicHeader = wpJObject["properties"]["topicHeader"].ToString();
                        }
                        if (wpJObject["properties"]["authors"] != null)
                        {
                            this.Authors = wpJObject["properties"]["authors"].ToString();
                        }
                        if (wpJObject["properties"]["authorByline"] != null)
                        {
                            this.AuthorByLine = wpJObject["properties"]["authorByline"].ToString();
                        }
#endif
                        // Specific properties that only apply when the header has a custom image
                        if (!string.IsNullOrEmpty(this.imageServerRelativeUrl))
                        {
                            this.pageHeaderType = ClientSidePageHeaderType.Custom;
                            if (wpJObject["properties"] != null)
                            {
                                Guid result = new Guid();
                                if (wpJObject["properties"]["siteId"] != null && Guid.TryParse(wpJObject["properties"]["siteId"].ToString(), out result))
                                {
                                    this.siteId = result;
                                }
                                if (wpJObject["properties"]["webId"] != null && Guid.TryParse(wpJObject["properties"]["webId"].ToString(), out result))
                                {
                                    this.webId = result;
                                }
                                if (wpJObject["properties"]["listId"] != null && Guid.TryParse(wpJObject["properties"]["listId"].ToString(), out result))
                                {
                                    this.listId = result;
                                }
                                if (wpJObject["properties"]["uniqueId"] != null && Guid.TryParse(wpJObject["properties"]["uniqueId"].ToString(), out result))
                                {
                                    this.uniqueId = result;
                                }

                                if (this.siteId != Guid.Empty && this.webId != Guid.Empty && this.listId != Guid.Empty && this.uniqueId != Guid.Empty)
                                {
                                    this.headerImageResolved = true;
                                }
                            }

                            System.Globalization.CultureInfo usCulture = new System.Globalization.CultureInfo("en-US");
                            System.Globalization.CultureInfo europeanCulture = new System.Globalization.CultureInfo("nl-BE");

                            if (wpJObject["properties"]["translateX"] != null)
                            {
                                double translateX = 0;
                                var translateXEN = wpJObject["properties"]["translateX"].ToString();

                                System.Globalization.CultureInfo cultureToUse;
                                if (translateXEN.Contains("."))
                                {
                                    cultureToUse = usCulture;
                                }
                                else if (translateXEN.Contains(","))
                                {
                                    cultureToUse = europeanCulture;
                                }
                                else
                                {
                                    cultureToUse = usCulture;
                                }

                                Double.TryParse(translateXEN, System.Globalization.NumberStyles.Float, cultureToUse, out translateX);
                                this.TranslateX = translateX;
                            }
                            if (wpJObject["properties"]["translateY"] != null)
                            {
                                double translateY = 0;
                                var translateYEN = wpJObject["properties"]["translateY"].ToString();

                                System.Globalization.CultureInfo cultureToUse;
                                if (translateYEN.Contains("."))
                                {
                                    cultureToUse = usCulture;
                                }
                                else if (translateYEN.Contains(","))
                                {
                                    cultureToUse = europeanCulture;
                                }
                                else
                                {
                                    cultureToUse = usCulture;
                                }

                                Double.TryParse(translateYEN, System.Globalization.NumberStyles.Float, cultureToUse, out translateY);
                                this.TranslateY = translateY;
                            }
#if !SP2019
                            if (wpJObject["properties"]["altText"] != null)
                            {
                                this.AlternativeText = wpJObject["properties"]["altText"].ToString();
                            }
#endif
                        }
                        else
                        {
                            this.pageHeaderType = ClientSidePageHeaderType.Default;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Returns the header html representation
        /// </summary>
        /// <param name="pageTitle">Title of the page</param>
        /// <returns>Header html value</returns>
        public string ToHtml(string pageTitle)
        {
            if (pageTitle == null)
            {
                pageTitle = "";
            }
            else
            {
                pageTitle = EncodePageTitle(pageTitle);
            }

            // Get the web part properties
            if (!string.IsNullOrEmpty(this.ImageServerRelativeUrl) && this.clientContext != null)
            {
                if (!headerImageResolved)
                {
                    ResolvePageHeaderImage();
                }

                if (headerImageResolved)
                {
                    string focalPoints = "";
                    if (TranslateX.HasValue || TranslateY.HasValue)
                    {
                        System.Globalization.CultureInfo usCulture = new System.Globalization.CultureInfo("en-US");
                        var translateX = TranslateX.Value.ToString(usCulture);
                        var translateY = TranslateY.Value.ToString(usCulture);
                        focalPoints = $",&quot;translateX&quot;&#58;{translateX},&quot;translateY&quot;&#58;{translateY}";
                    }

                    // Populate default properties
                    var header = FillDefaultProperties(CustomPageHeader);
                    // Populate custom header specific properties
#if !SP2019
                    return header.Replace("@@siteId@@", this.siteId.ToString()).Replace("@@webId@@", this.webId.ToString()).Replace("@@listId@@", this.listId.ToString()).Replace("@@uniqueId@@", this.uniqueId.ToString()).Replace("@@focalPoints@@", focalPoints).Replace("@@title@@", pageTitle).Replace("@@imageSource@@", this.ImageServerRelativeUrl).Replace("@@alternativetext@@", this.AlternativeText == null ? "" : this.AlternativeText);
#else
                    return header.Replace("@@siteId@@", this.siteId.ToString()).Replace("@@webId@@", this.webId.ToString()).Replace("@@listId@@", this.listId.ToString()).Replace("@@uniqueId@@", this.uniqueId.ToString()).Replace("@@focalPoints@@", focalPoints).Replace("@@title@@", pageTitle).Replace("@@imageSource@@", this.ImageServerRelativeUrl);
#endif
                }
            }

            // in case nothing worked out...
            // Populate default properties
            var defaultHeader = FillDefaultProperties(DefaultPageHeader);
            // Populate title
            return defaultHeader.Replace("@@title@@", pageTitle);
        }

        private string FillDefaultProperties(string header)
        {
#if !SP2019
            if (!string.IsNullOrEmpty(this.Authors))
            {
                string data = this.Authors.Replace("\r", "").Replace("\n", "").TrimStart(new char[] { '[' }).TrimEnd(new char[] { ']' });
                var jsonencoded = WebUtility.HtmlEncode(data).Replace(":", "&#58;"); //.Replace("@", "%40");
                header = header.Replace("@@authors@@", jsonencoded);
            }
            else
            {
                header = header.Replace("@@authors@@", "");
            }

            if (!string.IsNullOrEmpty(this.AuthorByLine))
            {
                string data = this.AuthorByLine.Replace("\r", "").Replace("\n", "").Replace(" ", "").TrimStart(new char[] { '[' }).TrimEnd(new char[] { ']' });
                var jsonencoded = WebUtility.HtmlEncode(data).Replace(":", "&#58;");
                header = header.Replace("@@authorbyline@@", jsonencoded);

                int userId = -1;
                try
                {
                    var user = this.clientContext.Web.EnsureUser(data.Replace("\"", "").Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[2]);
                    this.clientContext.Load(user);
                    this.clientContext.ExecuteQueryRetry();
                    userId = user.Id;
                }
                catch (Exception ex)
                {

                }

                this.AuthorByLineId = userId;
            }
            else
            {
                header = header.Replace("@@authorbyline@@", "");
            }

            return header.Replace("@@showtopicheader@@", this.ShowTopicHeader.ToString().ToLower()).Replace("@@showpublishdate@@", this.ShowPublishDate.ToString().ToLower()).Replace("@@topicheader@@", this.TopicHeader == null ? "" : this.TopicHeader).Replace("@@textalignment@@", this.TextAlignment.ToString()).Replace("@@layouttype@@", this.LayoutType.ToString());
#else
            return header.Replace("@@layouttype@@", this.LayoutType.ToString());
#endif
        }

        private static string Replace1point4Defaults(string header)
        {
            return header.Replace("@@showtopicheader@@", "false").Replace("@@showpublishdate@@", "false").Replace("@@topicheader@@", "");
        }

        private static string EncodePageTitle(string pageTitle)
        {
            string result = pageTitle;

            if (result.Contains("\""))
            {
                result = result.Replace("\"", "\\&quot;");
            }

            return result;
        }

        private void ResolvePageHeaderImage()
        {
            try
            {
                this.siteId = this.clientContext.Site.EnsureProperty(p => p.Id);
                this.webId = this.clientContext.Web.EnsureProperty(p => p.Id);

                if (!ImageServerRelativeUrl.StartsWith("/_LAYOUTS", StringComparison.OrdinalIgnoreCase))
                {
                    var pageHeaderImage = this.clientContext.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(ImageServerRelativeUrl));
                    this.clientContext.Load(pageHeaderImage, p => p.UniqueId, p => p.ListId);
                    this.clientContext.ExecuteQueryRetry();

                    this.listId = pageHeaderImage.ListId;
                    this.uniqueId = pageHeaderImage.UniqueId;
                }

                this.headerImageResolved = true;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    // provided file link does not exist...we're eating the exception and the page will end up with a default page header
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ClientSidePageHeader_ImageNotFound, ImageServerRelativeUrl);
                }
                else if (ex.Message.Contains("SPWeb.ServerRelativeUrl"))
                {
                    // image has to live in the web for which we've set up the client context...if not skip and log a warning
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ClientSidePageHeader_ImageInDifferentWeb, imageServerRelativeUrl);
                }
                else
                {
                    throw;
                }
            }
        }

    }
#endif
}
