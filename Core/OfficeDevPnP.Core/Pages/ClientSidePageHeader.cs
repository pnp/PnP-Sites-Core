using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    /// <summary>
    /// Class that implements the client side page header
    /// </summary>
    public class ClientSidePageHeader
    {
        private const string NoPageHeader      = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.3\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.3&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;0&#125;&#125;\"></div></div>";
        private const string DefaultPageHeader = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.3\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.3&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;4,&quot;translateX&quot;&#58;50,&quot;translateY&quot;&#58;50&#125;&#125;\"></div></div>";
        private const string CustomPageHeader  = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.3\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&quot;imageSource&quot;&#58;&quot;@@imageSource@@&quot;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.3&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;@@title@@&quot;,&quot;imageSourceType&quot;&#58;2,&quot;siteId&quot;&#58;&quot;@@siteId@@&quot;,&quot;webId&quot;&#58;&quot;@@webId@@&quot;,&quot;listId&quot;&#58;&quot;@@listId@@&quot;,&quot;uniqueId&quot;&#58;&quot;&#123;@@uniqueId@@&#125;&quot;@@focalPoints@@&#125;&#125;\"></div></div>";

        private string imageServerRelativeUrl;
        private ClientContext clientContext;
        private bool headerImageResolved = false;
        private Guid siteId;
        private Guid webId;
        private Guid listId;
        private Guid uniqueId;

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
        public string TranslateX { get; set; }

        /// <summary>
        /// Image focal point Y coordinate
        /// </summary>
        public string TranslateY { get; set; }

#region construction
        /// <summary>
        /// Creates a custom header with a custom image
        /// </summary>
        /// <param name="cc">ClientContext of the site hosting the image</param>
        /// <param name="imageServerRelativeUrl">Server relative image url</param>
        public ClientSidePageHeader(ClientContext cc, string imageServerRelativeUrl)
        {
            this.imageServerRelativeUrl = imageServerRelativeUrl;
            this.clientContext = cc;
        }

        /// <summary>
        /// Creates a custom header with a custom image + custom image offset
        /// </summary>
        /// <param name="cc">ClientContext of the site hosting the image</param>
        /// <param name="imageServerRelativeUrl">Server relative image url</param>
        /// <param name="translateX">X offset coordinate</param>
        /// <param name="translateY">Y offset coordinate</param>
        public ClientSidePageHeader(ClientContext cc, string imageServerRelativeUrl, string translateX, string translateY): this(cc, imageServerRelativeUrl)
        {
            TranslateX = translateX;
            TranslateY = translateY;
        }
#endregion

        /// <summary>
        /// Returns the header value to set a "no header" 
        /// </summary>
        /// <param name="pageTitle">Title of the page</param>
        /// <returns>Header html value that indicates "no header"</returns>
        public static string NoHeader(string pageTitle)
        {
            if (pageTitle == null)
            {
                pageTitle = "";
            }

            return NoPageHeader.Replace("@@title@@", pageTitle);
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
                    if (!string.IsNullOrEmpty(TranslateX) || !string.IsNullOrEmpty(TranslateY))
                    {
                        focalPoints = $",&quot;translateX&quot;&#58;{TranslateX},&quot;translateY&quot;&#58;{TranslateY}";
                    }

                    return CustomPageHeader.Replace("@@siteId@@", this.siteId.ToString()).Replace("@@webId@@", this.webId.ToString()).Replace("@@listId@@", this.listId.ToString()).Replace("@@uniqueId@@", this.uniqueId.ToString()).Replace("@@focalPoints@@", focalPoints).Replace("@@title@@", pageTitle); 
                }
            }

            // in case nothing worked out...
            return DefaultPageHeader.Replace("@@title@@", pageTitle);
        }

        private void ResolvePageHeaderImage()
        {
            try
            {
                this.clientContext.Site.EnsureProperties(p => p.Id);
                this.clientContext.Web.EnsureProperties(p => p.Id);

                var pageHeaderImage = this.clientContext.Web.GetFileByServerRelativeUrl(ImageServerRelativeUrl);
                this.clientContext.Load(pageHeaderImage, p => p.UniqueId, p => p.ListId);
                this.clientContext.ExecuteQueryRetry();

                this.siteId = this.clientContext.Site.Id;
                this.webId = this.clientContext.Web.Id;
                this.listId = pageHeaderImage.ListId;
                this.uniqueId = pageHeaderImage.UniqueId;
                this.headerImageResolved = true;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    // provided file link does not exist...we're eating the exception and the page will end up with a default page header
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.ClientSidePageHeader_ImageNotFound, ImageServerRelativeUrl);
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
