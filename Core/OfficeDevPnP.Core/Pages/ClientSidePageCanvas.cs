using AngleSharp.Parser.Html;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;

namespace OfficeDevPnP.Core.Pages
{
#if !ONPREMISES
    #region Canvas page model classes   
    /// <summary>
    /// List of possible OOB web parts
    /// </summary>
    public enum DefaultClientSideWebParts
    {
        /// <summary>
        /// Content Rollup webpart
        /// </summary>
        ContentRollup,
        /// <summary>
        /// Bing Map webpart
        /// </summary>
        BingMap,
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
        /// <summary>
        /// PowerBI Report Embed webpart
        /// </summary>
        PowerBIReportEmbed,
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
        /// <summary>
        /// Group Calendar webpart
        /// </summary>
        GroupCalendar,
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
        QuickLinks
    }

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

    /// <summary>
    /// Represents a modern client side page with all it's contents
    /// </summary>
    public class ClientSidePage
    {
        #region variables
        // fields
        public const string CanvasField = "CanvasContent1";
        public const string PageLayoutType = "PageLayoutType";
        public const string ApprovalStatus = "_ModerationStatus";
        public const string ContentTypeId = "ContentTypeId";
        public const string Title = "Title";
        public const string ClientSideApplicationId = "ClientSideApplicationId";
        public const string PromotedStateField = "PromotedState";
        public const string BannerImageUrl = "BannerImageUrl";
        public const string FirstPublishedDate = "FirstPublishedDate";
        public const string FileLeafRef = "FileLeafRef";

        // feature
        public const string SitePagesFeatureId = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec";

        private ClientContext context;
        private string pageName;
        private string pagesLibrary;
        private List spPagesLibrary;
        private ListItem pageListItem;
        private string sitePagesServerRelativeUrl;
        private bool securityInitialized = false;
        private string accessToken;
        private System.Collections.Generic.List<CanvasZone> zones = new System.Collections.Generic.List<CanvasZone>(1);
        private System.Collections.Generic.List<CanvasControl> controls = new System.Collections.Generic.List<CanvasControl>(5);
        private ClientSidePageLayoutType layoutType;
        private bool keepDefaultWebParts;
        private string pageTitle;
        #endregion

        #region construction
        /// <summary>
        /// Constructs ClientSidePage class
        /// </summary>
        /// <param name="clientSidePageLayoutType"><see cref="ClientSidePageLayoutType"/> type of the page to create. Defaults to Article type</param>
        public ClientSidePage(ClientSidePageLayoutType clientSidePageLayoutType = ClientSidePageLayoutType.Article)
        {
            this.layoutType = clientSidePageLayoutType;

            if (this.layoutType == ClientSidePageLayoutType.Home)
            {
                // By default we're assuming you want to have a customized home page, change this to true in case you want to create a home page holding the default OOB web parts
                this.keepDefaultWebParts = false;
            }

            //this.zones.Add(new CanvasZone(this, CanvasZoneTemplate.OneColumn, 0));
            this.pagesLibrary = "SitePages";
        }

        /// <summary>
        /// Constructs ClientSidePage class and connects a <see cref="ClientContext"/> instance, this is needed to allow interaction with SharePoint
        /// </summary>
        /// <param name="cc">The SharePoint <see cref="ClientContext"/> instance</param>
        /// <param name="clientSidePageLayoutType"><see cref="ClientSidePageLayoutType"/> type of the page to create. Defaults to Article type</param>
        public ClientSidePage(ClientContext cc, ClientSidePageLayoutType clientSidePageLayoutType = ClientSidePageLayoutType.Article) : this(clientSidePageLayoutType)
        {
            if (cc == null)
            {
                throw new ArgumentNullException("Passed ClientContext object cannot be null");
            }
            this.context = cc;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Title of the client side page
        /// </summary>
        public string PageTitle
        {
            get
            {
                return this.pageTitle;
            }
            set
            {
                this.pageTitle = value;
            }
        }

        /// <summary>
        /// Collection of Zones that exist on this client side page
        /// </summary>
        public System.Collections.Generic.List<CanvasZone> Zones
        {
            get
            {
                return this.zones;
            }
        }

        /// <summary>
        /// Collection of all control that exist on this client side page
        /// </summary>
        public System.Collections.Generic.List<CanvasControl> Controls
        {
            get
            {
                return this.controls;
            }
        }

        /// <summary>
        /// Layout type of the client side page
        /// </summary>
        public ClientSidePageLayoutType LayoutType
        {
            get
            {
                return this.layoutType;
            }
            set
            {
                this.layoutType = value;
            }
        }

        /// <summary>
        /// When a page of type Home is created you can opt to only keep the default client side web parts by setting this to true. This also is a way to reset your home page back the the stock one.
        /// </summary>
        public bool KeepDefaultWebParts
        {
            get
            {
                return this.keepDefaultWebParts;
            }
            set
            {
                this.keepDefaultWebParts = value;
            }
        }

        /// <summary>
        /// ClientContext object that will be used to read and write to SharePoint
        /// </summary>
        public ClientContext Context
        {
            get
            {
                return this.context;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("Passed ClientContext object cannot be null");
                }
                this.context = value;
            }
        }

        /// <summary>
        /// The site relative path to SitePages library
        /// </summary>
        public string PagesLibrary
        {
            get
            {
                return this.pagesLibrary;
            }
            set
            {
                if (String.IsNullOrEmpty(value))
                {
                    throw new ArgumentNullException("Passed pages library cannot be null or empty");
                }

                // validate the list existance in case we've a ClientContext object set
                if (this.Context != null)
                {
                    if (this.Context.Web.GetListByUrl(value) == null)
                    {
                        throw new ArgumentException("Passed pages library does not exist in current web");
                    }
                }

                this.pagesLibrary = value;
            }
        }

        /// <summary>
        /// The SharePoint list item of the saved/loaded page
        /// </summary>
        public ListItem PageListItem
        {
            get
            {
                return this.pageListItem;
            }
        }

        /// <summary>
        /// The default zone of the client side page
        /// </summary>
        public CanvasZone DefaultZone
        {
            get
            {
                // Add a default zone if there wasn't one yet created
                if (this.zones.Count == 0)
                {
                    this.zones.Add(new CanvasZone(this, CanvasZoneTemplate.OneColumn, 0));
                }

                return zones.First();
            }
        }

        /// <summary>
        /// Does this page have comments disabled
        /// </summary>
        public bool CommentsDisabled
        {
            get
            {
                EnsurePageListItem();
                if (this.PageListItem != null)
                {
                    this.PageListItem.EnsureProperty(p => p.CommentsDisabled);
                    return this.PageListItem.CommentsDisabled;
                }
                else
                {
                    throw new Exception("You first need to save the page before you check for CommentsEnabled status");
                }
            }
        }
        #endregion

        #region public methods

        /// <summary>
        /// Adds a new zone to your client side page
        /// </summary>
        /// <param name="template">The <see cref="CanvasZoneTemplate"/> type of the zone</param>
        /// <param name="order">Controls the order of the new zone</param>
        public void AddZone(CanvasZoneTemplate template, int order)
        {
            var zone = new CanvasZone(this, template, order);
            AddZone(zone);
        }

        /// <summary>
        /// Adds a new zone to your client side page
        /// </summary>
        /// <param name="zone"><see cref="CanvasZone"/> object describing the zone to add</param>
        public void AddZone(CanvasZone zone)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }
            this.zones.Add(zone);
        }

        /// <summary>
        /// Adds a new zone to your client side page with a given order
        /// </summary>
        /// <param name="zone"><see cref="CanvasZone"/> object describing the zone to add</param>
        /// <param name="order">Controls the order of the new zone</param>
        public void AddZone(CanvasZone zone, float order)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }
            zone.Order = order;
            this.zones.Add(zone);
        }

        /// <summary>
        /// Adds a new control to your client side page using the default <see cref="CanvasZone"/>
        /// </summary>
        /// <param name="control"><see cref="CanvasControl"/> to add</param>
        public void AddControl(CanvasControl control)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }

            // add to defaultzone and section
            if (control.Zone == null)
            {
                control.zone = this.DefaultZone;
            }
            if (control.Section == null)
            {
                control.section = this.DefaultZone.DefaultSection;
            }

            this.controls.Add(control);
        }

        /// <summary>
        /// Adds a new control to your client side page using the default <see cref="CanvasZone"/> using a given order
        /// </summary>
        /// <param name="control"><see cref="CanvasControl"/> to add</param>
        /// <param name="order">Order of the control in the default zone</param>
        public void AddControl(CanvasControl control, int order)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }

            // add to defaultzone and section
            if (control.Zone == null)
            {
                control.zone = this.DefaultZone;
            }
            if (control.Section == null)
            {
                control.section = this.DefaultZone.DefaultSection;
            }
            control.Order = order;

            this.controls.Add(control);
        }

        /// <summary>
        /// Adds a new control to your client side page in the given zone
        /// </summary>
        /// <param name="control"><see cref="CanvasControl"/> to add</param>
        /// <param name="zone"><see cref="CanvasZone"/> that will hold the control. Control will end up in the <see cref="CanvasZone.DefaultSection"/>.</param>
        public void AddControl(CanvasControl control, CanvasZone zone)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }

            control.zone = zone;
            control.section = zone.DefaultSection;

            this.controls.Add(control);
        }

        /// <summary>
        /// Adds a new control to your client side page in the given zone with a given order
        /// </summary>
        /// <param name="control"><see cref="CanvasControl"/> to add</param>
        /// <param name="zone"><see cref="CanvasZone"/> that will hold the control. Control will end up in the <see cref="CanvasZone.DefaultSection"/>.</param>
        /// <param name="order">Order of the control in the given zone</param>
        public void AddControl(CanvasControl control, CanvasZone zone, int order)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }

            control.zone = zone;
            control.section = zone.DefaultSection;
            control.Order = order;

            this.controls.Add(control);
        }

        /// <summary>
        /// Adds a new control to your client side page in the given section
        /// </summary>
        /// <param name="control"><see cref="CanvasControl"/> to add</param>
        /// <param name="section"><see cref="CanvasSection"/> that will hold the control</param>    
        public void AddControl(CanvasControl control, CanvasSection section)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            control.zone = section.Zone;
            control.section = section;

            this.controls.Add(control);
        }

        /// <summary>
        /// Adds a new control to your client side page in the given section with a given order
        /// </summary>
        /// <param name="control"><see cref="CanvasControl"/> to add</param>
        /// <param name="section"><see cref="CanvasSection"/> that will hold the control</param>    
        /// <param name="order">Order of the control in the given section</param>
        public void AddControl(CanvasControl control, CanvasSection section, int order)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            control.zone = section.Zone;
            control.section = section;
            control.Order = order;

            this.controls.Add(control);
        }

        /// <summary>
        /// Deletes a control from a page
        /// </summary>
        public void Delete()
        {
            if (this.pageListItem == null)
            {
                throw new ArgumentException($"Page {this.pageName} was not loaded/saved to SharePoint and therefore can't be deleted");
            }

            pageListItem.DeleteObject();
            this.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Returns the html representation of this client side page. This is the content that will be persisted in the <see cref="ClientSidePage.PageListItem"/> list item.
        /// </summary>
        /// <returns>Html representation</returns>
        public string ToHtml()
        {
            StringBuilder html = new StringBuilder(100);
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;

                htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);

                foreach (var zone in this.zones.OrderBy(p => p.Order))
                {
                    htmlWriter.Write(zone.ToHtml());
                }

                htmlWriter.RenderEndTag();
            }

            return html.ToString();
        }

        /// <summary>
        /// Loads an existint SharePoint client side page
        /// </summary>
        /// <param name="cc">ClientContext object used to load the page</param>
        /// <param name="pageName">Name of the page (e.g. mypage.aspx) to load</param>
        /// <returns>A <see cref="ClientSidePage"/> instance for the given page</returns>
        public static ClientSidePage Load(ClientContext cc, string pageName)
        {
            if (cc == null)
            {
                throw new ArgumentNullException("Passed ClientContext object cannot be null");
            }

            if (String.IsNullOrEmpty(pageName))
            {
                throw new ArgumentException("Passed pageName object cannot be null or empty");
            }

            ClientSidePage page = new ClientSidePage(cc)
            {
                pageName = pageName
            };

            var pagesLibrary = page.Context.Web.GetListByUrl(page.PagesLibrary, p => p.RootFolder);
            page.sitePagesServerRelativeUrl = pagesLibrary.RootFolder.ServerRelativeUrl;

            var file = page.Context.Web.GetFileByServerRelativeUrl($"{page.sitePagesServerRelativeUrl}/{page.pageName}");
            page.Context.Web.Context.Load(file, f => f.ListItemAllFields, f => f.Exists);
            page.Context.Web.Context.ExecuteQueryRetry();

            if (!file.Exists)
            {
                throw new ArgumentException($"Page {pageName} does not exist in current web");
            }

            var item = file.ListItemAllFields;
            page.LayoutType = (ClientSidePageLayoutType)Enum.Parse(typeof(ClientSidePageLayoutType), item[ClientSidePage.PageLayoutType].ToString());
            if (!(item[ClientSidePage.CanvasField] == null || string.IsNullOrEmpty(item[ClientSidePage.CanvasField].ToString())))
            {
                var html = item[ClientSidePage.CanvasField].ToString();

                if (string.IsNullOrEmpty(html))
                {
                    throw new ArgumentException($"Page {pageName} is not a \"modern\" client side page");
                }

                page.pageListItem = item;
                page.LoadFromHtml(html);
            }
            return page;
        }

        /// <summary>
        /// Persists the current <see cref="ClientSidePage"/> instance as a client side page in SharePoint
        /// </summary>
        /// <param name="pageName">Name of the page (e.g. mypage.aspx) to save</param>
        public void Save(string pageName = null)
        {
            string serverRelativePageName;
            File pageFile;
            ListItem item;

            // Validate we're not using "wrong" layouts for the given site type
            ValidateOneColumnFullWidthZoneUsage();

            // Try to load the page
            LoadPageFile(pageName, out serverRelativePageName, out pageFile);

            if (!pageFile.Exists)
            {
                // create page listitem
                item = this.spPagesLibrary.RootFolder.Files.AddTemplateFile(serverRelativePageName, TemplateFileType.ClientSidePage).ListItemAllFields;
                // Fix page to be modern
                item[ClientSidePage.ContentTypeId] = BuiltInContentTypeId.ModernArticlePage;
                item[ClientSidePage.Title] = string.IsNullOrWhiteSpace(this.pageTitle) ? System.IO.Path.GetFileNameWithoutExtension(this.pageName) : pageTitle;
                item[ClientSidePage.ClientSideApplicationId] = ClientSidePage.SitePagesFeatureId;
                item[ClientSidePage.PageLayoutType] = this.layoutType.ToString();
                if (this.layoutType == ClientSidePageLayoutType.Article)
                {
                    item[ClientSidePage.PromotedStateField] = (Int32)PromotedState.NotPromoted;
                    item[ClientSidePage.BannerImageUrl] = "/_layouts/15/images/sitepagethumbnail.png";
                }
                item.Update();
                this.Context.Web.Context.Load(item);
                this.Context.Web.Context.ExecuteQueryRetry();
            }
            else
            {
                item = pageFile.ListItemAllFields;
            }

            // Persist to page field
            if (this.layoutType == ClientSidePageLayoutType.Home && this.KeepDefaultWebParts)
            {
                item[ClientSidePage.CanvasField] = "";
            }
            else
            {
                item[ClientSidePage.CanvasField] = this.ToHtml();
            }
            item.Update();
            this.Context.ExecuteQueryRetry();

            this.pageListItem = item;
        }

        /// <summary>
        /// Instantiate a <see cref="ClientSidePage"/> from a html fragment
        /// </summary>
        /// <param name="html">Html to convert into a <see cref="ClientSidePage"/></param>
        /// <returns>A <see cref="ClientSidePage"/> instance</returns>
        public static ClientSidePage FromHtml(string html)
        {
            if (String.IsNullOrEmpty(html))
            {
                throw new ArgumentException("Passed html cannot be null or empty");
            }

            ClientSidePage page = new ClientSidePage();
            page.LoadFromHtml(html);
            return page;
        }

        /// <summary>
        /// Return the name (=guid) for a given first party out of the box web part
        /// </summary>
        /// <param name="webPart">First party web part</param>
        /// <returns>Name(=guid) for the given web part</returns>
        public static string ClientSideWebPartEnumToName(DefaultClientSideWebParts webPart)
        {
            switch (webPart)
            {
                case DefaultClientSideWebParts.ContentRollup: return "daf0b71c-6de8-4ef7-b511-faae7c388708";
                case DefaultClientSideWebParts.BingMap: return "e377ea37-9047-43b9-8cdb-a761be2f8e09";
                case DefaultClientSideWebParts.ContentEmbed: return "490d7c76-1824-45b2-9de3-676421c997fa";
                case DefaultClientSideWebParts.DocumentEmbed: return "b7dd04e1-19ce-4b24-9132-b60a1c2b910d";
                case DefaultClientSideWebParts.Image: return "d1d91016-032f-456d-98a4-721247c305e8";
                case DefaultClientSideWebParts.ImageGallery: return "af8be689-990e-492a-81f7-ba3e4cd3ed9c";
                case DefaultClientSideWebParts.LinkPreview: return "6410b3b6-d440-4663-8744-378976dc041e";
                case DefaultClientSideWebParts.NewsFeed: return "0ef418ba-5d19-4ade-9db0-b339873291d0";
                case DefaultClientSideWebParts.NewsReel: return "a5df8fdf-b508-4b66-98a6-d83bc2597f63";
                case DefaultClientSideWebParts.PowerBIReportEmbed: return "58fcd18b-e1af-4b0a-b23b-422c2c52d5a2";
                case DefaultClientSideWebParts.QuickChart: return "91a50c94-865f-4f5c-8b4e-e49659e69772";
                case DefaultClientSideWebParts.SiteActivity: return "eb95c819-ab8f-4689-bd03-0c2d65d47b1f";
                case DefaultClientSideWebParts.VideoEmbed: return "275c0095-a77e-4f6d-a2a0-6a7626911518";
                case DefaultClientSideWebParts.YammerEmbed: return "31e9537e-f9dc-40a4-8834-0e3b7df418bc";
                case DefaultClientSideWebParts.Events: return "20745d7d-8581-4a6c-bf26-68279bc123fc";
                case DefaultClientSideWebParts.GroupCalendar: return "6676088b-e28e-4a90-b9cb-d0d0303cd2eb";
                case DefaultClientSideWebParts.Hero: return "c4bd7b2f-7b6e-4599-8485-16504575f590";
                case DefaultClientSideWebParts.List: return "f92bf067-bc19-489e-a556-7fe95f508720";
                case DefaultClientSideWebParts.PageTitle: return "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788";
                case DefaultClientSideWebParts.People: return "7f718435-ee4d-431c-bdbf-9c4ff326f46e";
                case DefaultClientSideWebParts.QuickLinks: return "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd";
                default: return "";
            }
        }

        /// <summary>
        /// Creates an instance of an out of the box (default, first party) client side web part
        /// </summary>
        /// <param name="webPart">The out of the box web part you want to instantiate</param>
        /// <returns><see cref="ClientSideWebPart"/> instance</returns>
        public ClientSideWebPart InstantiateDefaultWebPart(DefaultClientSideWebParts webPart)
        {
            var webPartName = ClientSidePage.ClientSideWebPartEnumToName(webPart);
            var webParts = this.AvailableClientSideComponents(webPartName);

            if (webParts.Count() == 1)
            {
                return new ClientSideWebPart(webParts.First());
            }

            return null;
        }

        /// <summary>
        /// Gets a list of available client side web parts to use
        /// </summary>
        /// <returns>List of available <see cref="ClientSideComponent"/></returns>
        public System.Collections.Generic.IEnumerable<ClientSideComponent> AvailableClientSideComponents()
        {
            return this.AvailableClientSideComponents(null);
        }

        /// <summary>
        /// Gets an out of the box, default, client side web parts to use
        /// </summary>
        /// <param name="webPart">Get one of the default, out of the box client side web parts</param>
        /// <returns>List of available <see cref="ClientSideComponent"/></returns>
        public System.Collections.Generic.IEnumerable<ClientSideComponent> AvailableClientSideComponents(DefaultClientSideWebParts webPart)
        {
            return this.AvailableClientSideComponents(ClientSidePage.ClientSideWebPartEnumToName(webPart));
        }

        /// <summary>
        /// Gets a list of available client side web parts to use having a given name
        /// </summary>
        /// <param name="name">Name of the web part to retrieve</param>
        /// <returns>List of available <see cref="ClientSideComponent"/></returns>
        public System.Collections.Generic.IEnumerable<ClientSideComponent> AvailableClientSideComponents(string name)
        {
            if (!this.securityInitialized)
            {
                this.InitializeSecurity();
            }

            // Request information about the available client side components from SharePoint
            Task<String> availableClientSideComponentsJson = Task.WhenAny(
                GetClientSideWebPartsAsync(this.accessToken, this.Context)
                ).Result;

            if (String.IsNullOrEmpty(availableClientSideComponentsJson.Result))
            {
                throw new ArgumentException("No client side components could be returned for this web...should not happen but it did...");
            }

            // Deserialize the returned data
            var jsonSerializerSettings = new JsonSerializerSettings();
            jsonSerializerSettings.MissingMemberHandling = MissingMemberHandling.Ignore;

            var clientSideComponents = ((System.Collections.Generic.IEnumerable<ClientSideComponent>)JsonConvert.DeserializeObject<AvailableClientSideComponents>(availableClientSideComponentsJson.Result, jsonSerializerSettings).value);

            if (clientSideComponents.Count() == 0)
            {
                throw new ArgumentException("No client side components could be returned for this web...should not happen but it did...");
            }

            if (!String.IsNullOrEmpty(name))
            {
                return clientSideComponents.Where(p => p.Name == name);
            }

            return clientSideComponents;
        }

        /// <summary>
        /// Publishes a client side page
        /// </summary>
        public void Publish()
        {
            Publish("");
        }

        /// <summary>
        /// Publishes a client side page
        /// </summary>
        /// <param name="publishMessage">Publish message</param>
        public void Publish(string publishMessage)
        {
            // Load the page
            string serverRelativePageName;
            File pageFile;

            LoadPageFile(pageName, out serverRelativePageName, out pageFile);

            if (pageFile.Exists)
            {
                // connect up the page list item for future reference
                this.pageListItem = pageFile.ListItemAllFields;
                // publish the page
                pageFile.Publish(publishMessage);
                this.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Enable commenting on this page
        /// </summary>
        public void EnableComments()
        {
            EnableCommentsImplementation(true);
        }

        /// <summary>
        /// Disable commenting on this page
        /// </summary>
        public void DisableComments()
        {
            EnableCommentsImplementation(false);
        }

        /// <summary>
        /// Demotes an client side <see cref="ClientSidePageLayoutType.Article"/> news page as a regular client side page
        /// </summary>
        public void DemoteNewsArticle()
        {
            if (this.LayoutType != ClientSidePageLayoutType.Article)
            {
                throw new Exception("You can't promote a home page as news article");
            }

            // ensure we do have the page list item loaded
            EnsurePageListItem();

            // Set promoted state
            this.pageListItem[ClientSidePage.PromotedStateField] = (Int32)PromotedState.NotPromoted;
            this.pageListItem.Update();
            this.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Promotes a regular <see cref="ClientSidePageLayoutType.Article"/> client side page as a news page
        /// </summary>
        public void PromoteAsNewsArticle()
        {
            if (this.LayoutType != ClientSidePageLayoutType.Article)
            {
                throw new Exception("You can only promote article pages as news article");
            }

            // ensure we do have the page list item loaded
            EnsurePageListItem();

            // Set promoted state
            this.pageListItem[ClientSidePage.PromotedStateField] = (Int32)PromotedState.Promoted;
            // Set publication date
            this.pageListItem[ClientSidePage.FirstPublishedDate] = DateTime.UtcNow;
            this.pageListItem.Update();
            this.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Sets the current <see cref="ClientSidePage"/> as home page for the current site
        /// </summary>
        public void PromoteAsHomePage()
        {
            if (this.LayoutType != ClientSidePageLayoutType.Home)
            {
                throw new Exception("You can only promote home pages as site home page");
            }

            // ensure we do have the page list item loaded
            EnsurePageListItem();

            this.Context.Web.EnsureProperty(p => p.RootFolder);
            this.Context.Web.RootFolder.WelcomePage = $"{this.PagesLibrary}/{this.PageListItem[ClientSidePage.FileLeafRef].ToString()}";
            this.Context.ExecuteQueryRetry();
        }
        #endregion

        #region Internal and private methods
        private void EnableCommentsImplementation(bool enable)
        {
            // ensure we do have the page list item loaded
            EnsurePageListItem();
            if (this.PageListItem != null)
            {
                this.pageListItem.SetCommentsDisabled(!enable);
                this.Context.ExecuteQueryRetry();
            }
            else
            {
                throw new Exception("This page first needs to be saved before comments can be enabled or disabled");
            }
        }

        private void ValidateOneColumnFullWidthZoneUsage()
        {
            bool hasOneColumnFullWidthZone = false;
            foreach (var zone in this.zones)
            {
                if (zone.Type == CanvasZoneTemplate.OneColumnFullWidth)
                {
                    hasOneColumnFullWidthZone = true;
                    break;
                }
            }
            if (hasOneColumnFullWidthZone)
            {
                this.Context.Web.EnsureProperties(p => p.WebTemplate, p => p.Configuration);
                if (!this.Context.Web.WebTemplate.Equals("SITEPAGEPUBLISHING", StringComparison.InvariantCultureIgnoreCase))
                {
                    throw new Exception($"You can't use a OneColumnFullWidth zone in this site template ({this.Context.Web.WebTemplate})");
                }
            }
        }

        private void EnsurePageListItem()
        {
            if (this.pageListItem == null)
            {
                string serverRelativePageName;
                File pageFile;
                LoadPageFile(this.pageName, out serverRelativePageName, out pageFile);
                if (pageFile.Exists)
                {
                    // connect up the page list item for future reference
                    this.pageListItem = pageFile.ListItemAllFields;
                }
            }
        }

        private void LoadPageFile(string pageName, out string serverRelativePageName, out File pageFile)
        {
            // Save page contents to SharePoint
            if (this.Context == null)
            {
                throw new Exception("No valid ClientContext object connected, can't save this page to SharePoint");
            }

            // Grab pages library reference
            if (this.spPagesLibrary == null)
            {
                this.spPagesLibrary = this.Context.Web.GetListByUrl(this.PagesLibrary, p => p.RootFolder);
            }

            // Build up server relative page url
            if (string.IsNullOrEmpty(this.sitePagesServerRelativeUrl))
            {
                this.sitePagesServerRelativeUrl = this.spPagesLibrary.RootFolder.ServerRelativeUrl;
            }

            if (!String.IsNullOrWhiteSpace(pageName))
            {
                this.pageName = pageName;
            }

            if (string.IsNullOrWhiteSpace(this.pageName))
            {
                throw new Exception("No valid page name specified, can't save this page to SharePoint");
            }

            serverRelativePageName = $"{this.sitePagesServerRelativeUrl}/{this.pageName}";

            // ensure page exists
            pageFile = this.Context.Web.GetFileByServerRelativeUrl(serverRelativePageName);
            this.Context.Web.Context.Load(pageFile, f => f.ListItemAllFields, f => f.Exists);
            this.Context.Web.Context.ExecuteQueryRetry();
        }

        private void LoadFromHtml(string html)
        {
            if (String.IsNullOrEmpty(html))
            {
                throw new ArgumentException("Passed html cannot be null or empty");
            }

            HtmlParser parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true });
            using (var document = parser.Parse(html))
            {
                // select all control div's
                var clientSideControls = document.All.Where(m => m.HasAttribute(CanvasControl.ControlDataAttribute));

                // clear zones as we're constructing them from the loaded html
                this.zones.Clear();

                int controlOrder = 0;
                foreach (var clientSideControl in clientSideControls)
                {
                    var controlData = clientSideControl.GetAttribute(CanvasControl.ControlDataAttribute);
                    var controlType = CanvasControl.GetType(controlData);

                    if (controlType == typeof(ClientSideText))
                    {
                        var control = new ClientSideText()
                        {
                            Order = controlOrder
                        };
                        control.FromHtml(clientSideControl);

                        // Handle control positioning in zones and sections
                        ApplyZoneAndSection(control, control.SpControlData.Position);

                        this.AddControl(control);
                    }
                    else if (controlType == typeof(ClientSideWebPart))
                    {
                        var control = new ClientSideWebPart();
                        control.FromHtml(clientSideControl);

                        // Handle control positioning in zones and sections
                        ApplyZoneAndSection(control, control.SpControlData.Position);

                        this.AddControl(control);
                    }
                    else if (controlType == typeof(CanvasSection))
                    {
                        var jsonSerializerSettings = new JsonSerializerSettings()
                        {
                            MissingMemberHandling = MissingMemberHandling.Ignore
                        };
                        var sectionData = JsonConvert.DeserializeObject<ClientSideCanvasData>(controlData, jsonSerializerSettings);
                        
                        var currentZone = this.zones.Where(p => p.Order == sectionData.Position.ZoneIndex).FirstOrDefault();
                        if (currentZone == null)
                        {
                            this.AddZone(new CanvasZone(this), sectionData.Position.ZoneIndex);
                            currentZone = this.zones.Where(p => p.Order == sectionData.Position.ZoneIndex).First();
                        }

                        var currentSection = currentZone.Sections.Where(p => p.Order == sectionData.Position.SectionIndex).FirstOrDefault();
                        if (currentSection == null)
                        {
                            CanvasSection newSection = new CanvasSection(currentZone);
                            currentZone.AddSection(new CanvasSection(currentZone, sectionData.Position.SectionIndex, sectionData.Position.SectionFactor));
                            currentSection = currentZone.Sections.Where(p => p.Order == sectionData.Position.SectionIndex).First();
                        }
                    }

                    controlOrder++;
                }
            }

            // Perform zone type detection
            foreach(var zone in this.zones)
            {
                if (zone.Sections.Count == 1)
                {
                    if (zone.Sections[0].SectionFactor == 0)
                    {
                        zone.Type = CanvasZoneTemplate.OneColumnFullWidth;
                    }
                    else
                    {
                        zone.Type = CanvasZoneTemplate.OneColumn;
                    }
                }
                else if (zone.Sections.Count == 2)
                {
                    if (zone.Sections[0].SectionFactor == 6)
                    {
                        zone.Type = CanvasZoneTemplate.TwoColumn;
                    }
                    else if (zone.Sections[0].SectionFactor == 4)
                    {
                        zone.Type = CanvasZoneTemplate.TwoColumnRight;
                    }
                    else if (zone.Sections[0].SectionFactor == 8)
                    {
                        zone.Type = CanvasZoneTemplate.TwoColumnLeft;
                    }
                }
                else if (zone.Sections.Count == 3)
                {
                    zone.Type = CanvasZoneTemplate.ThreeColumn;
                }
            }

        }

        private void ApplyZoneAndSection(CanvasControl control, ClientSideCanvasControlPosition position)
        {
            var currentZone = this.zones.Where(p => p.Order == position.ZoneIndex).FirstOrDefault();
            if (currentZone == null)
            {
                this.AddZone(new CanvasZone(this), position.ZoneIndex);
                currentZone = this.zones.Where(p => p.Order == position.ZoneIndex).First();
            }

            var currentSection = currentZone.Sections.Where(p => p.Order == position.SectionIndex).FirstOrDefault();
            if (currentSection == null)
            {
                CanvasSection newSection = new CanvasSection(currentZone);
                currentZone.AddSection(new CanvasSection(currentZone, position.SectionIndex, position.SectionFactor));
                currentSection = currentZone.Sections.Where(p => p.Order == position.SectionIndex).First();
            }

            control.zone = currentZone;
            control.section = currentSection;
        }

        private async Task<string> GetClientSideWebPartsAsync(string accessToken, ClientContext context)
        {
            string responseString = null;

            using (var handler = new HttpClientHandler())
            {
                // we're not in app-only or user + app context, so let's fall back to cookie based auth
                if (String.IsNullOrEmpty(accessToken))
                {
                    handler.Credentials = context.Credentials;
                    handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));
                }

                using (var httpClient = new HttpClient(handler))
                {
                    //GET https://bertonline.sharepoint.com/sites/130023/_api/web/GetClientSideWebParts HTTP/1.1

                    string requestUrl = String.Format("{0}/_api/web/GetClientSideWebParts", context.Web.Url);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("accept", "application/json;odata.metadata=minimal");
                    request.Headers.Add("odata-version", "4.0");

                    // We've an access token, so we're in app-only or user + app context
                    if (!String.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    }

                    HttpResponseMessage response = await httpClient.SendAsync(request);

                    if (response.IsSuccessStatusCode)
                    {
                        responseString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        // Something went wrong...
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
                return await Task.Run(() => responseString);
            }
        }

        private void InitializeSecurity()
        {
            // Let's try to grab an access token, will work when we're in app-only or user+app model
            this.Context.ExecutingWebRequest += Context_ExecutingWebRequest;
            this.Context.Load(this.Context.Web, w => w.Url);
            this.context.ExecuteQueryRetry();
            this.Context.ExecutingWebRequest -= Context_ExecutingWebRequest;
        }

        private void Context_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            if (!String.IsNullOrEmpty(e.WebRequestExecutor.RequestHeaders.Get("Authorization")))
            {
                this.accessToken = e.WebRequestExecutor.RequestHeaders.Get("Authorization").Replace("Bearer ", "");
            }
        }
        #endregion
    }

    /// <summary>
    /// The type of canvas being used
    /// </summary>
    public enum CanvasZoneTemplate
    {
        /// <summary>
        /// One column
        /// </summary>
        OneColumn = 0,
        /// <summary>
        /// One column, full browser width. This one only works for communication sites in combination with image or hero webparts
        /// </summary>
        OneColumnFullWidth =1,
        /// <summary>
        /// Two columns of the same size
        /// </summary>
        TwoColumn = 2,
        /// <summary>
        /// Three columns of the same size
        /// </summary>
        ThreeColumn = 3,
        /// <summary>
        /// Two columns, left one is 2/3, right one 1/3
        /// </summary>
        TwoColumnLeft = 4,
        /// <summary>
        /// Two columns, left one is 1/3, right one 2/3
        /// </summary>
        TwoColumnRight = 5,

    }

    /// <summary>
    /// Represents a zone on the canvas
    /// </summary>
    public class CanvasZone
    {
        #region variables
        private System.Collections.Generic.List<CanvasSection> sections = new System.Collections.Generic.List<CanvasSection>(3);
        private ClientSidePage page;
        #endregion

        #region construction
        internal CanvasZone(ClientSidePage page)
        {
            if (page == null)
            {
                throw new ArgumentNullException("Passed page cannot be null");
            }

            this.page = page;
            Order = 0;
        }

        /// <summary>
        /// Creates a new canvas zone
        /// </summary>
        /// <param name="page"><see cref="ClientSidePage"/> instance that holds this zone</param>
        /// <param name="canvasSectionTemplate">Type of zone to create</param>
        /// <param name="order">Order of this zone in the collection of zones on the page</param>
        public CanvasZone(ClientSidePage page, CanvasZoneTemplate canvasSectionTemplate, int order)
        {
            if (page == null)
            {
                throw new ArgumentNullException("Passed page cannot be null");
            }

            this.page = page;
            Type = canvasSectionTemplate;
            Order = order;

            switch (canvasSectionTemplate)
            {
                case CanvasZoneTemplate.OneColumn:
                    goto default;
                case CanvasZoneTemplate.OneColumnFullWidth:
                    this.sections.Add(new CanvasSection(this, 1, 0));
                    break;
                case CanvasZoneTemplate.TwoColumn:
                    this.sections.Add(new CanvasSection(this, 1, 6));
                    this.sections.Add(new CanvasSection(this, 2, 6));
                    break;
                case CanvasZoneTemplate.ThreeColumn:
                    this.sections.Add(new CanvasSection(this, 1, 4));
                    this.sections.Add(new CanvasSection(this, 2, 4));
                    this.sections.Add(new CanvasSection(this, 3, 4));
                    break;
                case CanvasZoneTemplate.TwoColumnLeft:
                    this.sections.Add(new CanvasSection(this, 1, 8));
                    this.sections.Add(new CanvasSection(this, 2, 4));
                    break;
                case CanvasZoneTemplate.TwoColumnRight:
                    this.sections.Add(new CanvasSection(this, 1, 4));
                    this.sections.Add(new CanvasSection(this, 2, 8));
                    break;
                default:
                    this.sections.Add(new CanvasSection(this, 1, 12));
                    break;
            }            
        }
        #endregion

        #region Properties
        /// <summary>
        /// Type of the zone
        /// </summary>
        public CanvasZoneTemplate Type { get; set; }

        /// <summary>
        /// Order in which this zone is presented on the page
        /// </summary>
        public float Order { get; set; }

        /// <summary>
        /// <see cref="CanvasSection"/> instances that are part of this zone
        /// </summary>
        public System.Collections.Generic.List<CanvasSection> Sections
        {
            get
            {
                return this.sections;
            }
        }

        /// <summary>
        /// The <see cref="ClientSidePage"/> instance holding this zone
        /// </summary>
        public ClientSidePage Page
        {
            get
            {
                return this.page;
            }
        }

        /// <summary>
        /// Controls hosted in this zone
        /// </summary>
        public System.Collections.Generic.List<CanvasControl> Controls
        {
            get
            {
                return this.Page.Controls.Where(p => p.Zone == this).ToList<CanvasControl>();
            }
        }

        /// <summary>
        /// The default <see cref="CanvasSection"/> of this zone
        /// </summary>
        public CanvasSection DefaultSection
        {
            get
            {
                if (this.sections.Count == 0)
                {
                    this.sections.Add(new CanvasSection(this));
                }

                return this.sections.First();
            }
        }
        #endregion

        #region public methods
        /// <summary>
        /// Renders this zone as a HTML fragment
        /// </summary>
        /// <returns>HTML string representing this zone</returns>
        public string ToHtml()
        {
            StringBuilder html = new StringBuilder(100);
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;

                foreach (var section in this.sections.OrderBy(z => z.Order))
                {
                    htmlWriter.Write(section.ToHtml());
                }
            }

            return html.ToString();
        }
        #endregion

        #region internal and private methods
        internal void AddSection(CanvasSection section)
        {
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            this.sections.Add(section);
        }
        #endregion
    }

    /// <summary>
    /// Represents a section in a canvas zone
    /// </summary>
    public class CanvasSection
    {
        #region variables
        public const string CanvasControlAttribute = "data-sp-canvascontrol";
        public const string CanvasDataVersionAttribute = "data-sp-canvasdataversion";
        public const string ControlDataAttribute = "data-sp-controldata";

        private int sectionFactor;
        private CanvasZone zone;
        private string DataVersion = "1.0";
        #endregion

        // internal constructors as we don't want users to manually create sections
        #region construction
        internal CanvasSection(CanvasZone zone)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }

            this.zone = zone;
            this.sectionFactor = 12;
            this.Order = 0;
        }

        internal CanvasSection(CanvasZone zone, int order)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }

            this.zone = zone;
            this.Order = order;
        }

        internal CanvasSection(CanvasZone zone, int order, int sectionFactor)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }

            this.zone = zone;
            this.Order = order;
            this.sectionFactor = sectionFactor;
        }
        #endregion

        #region Properties
        internal int Order { get; set; }

        /// <summary>
        /// <see cref="CanvasZone"/> this section belongs to
        /// </summary>
        public CanvasZone Zone
        {
            get
            {
                return this.zone;
            }
        }

        /// <summary>
        /// Section size factor. Max value is 12 (= one column), other options are 8,6,4 or 0
        /// </summary>
        public int SectionFactor
        {
            get
            {
                return this.sectionFactor;
            }
        }

        /// <summary>
        /// List of <see cref="CanvasControl"/> instances that are hosted in this section
        /// </summary>
        public System.Collections.Generic.List<CanvasControl> Controls
        {
            get
            {
                return this.Zone.Page.Controls.Where(p => p.Zone == this.Zone && p.Section == this).ToList<CanvasControl>();
            }
        }
        #endregion

        #region public methods
        /// <summary>
        /// Renders a HTML presentation of this section
        /// </summary>
        /// <returns>The HTML presentation of this section</returns>
        public string ToHtml()
        {
            StringBuilder html = new StringBuilder(100);
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;

                bool controlWrittenToSection = false;
                int controlIndex = 0;
                foreach (var control in this.Zone.Page.Controls.Where(p => p.Zone == this.Zone && p.Section == this).OrderBy(z => z.Order))
                {
                    controlIndex++;
                    htmlWriter.Write(control.ToHtml(controlIndex));
                    controlWrittenToSection = true;
                }

                // if a section does not contain a control we still need to render it, otherwise it get's "lost"
                if (!controlWrittenToSection)
                {
                    // Obtain the json data
                    var clientSideCanvasPosition = new ClientSideCanvasData()
                    {
                        Position = new ClientSideCanvasPosition()
                        {
                            ZoneIndex = this.Zone.Order,
                            SectionIndex = this.Order,
                            SectionFactor = this.SectionFactor,
                        }
                    };

                    var jsonControlData = JsonConvert.SerializeObject(clientSideCanvasPosition);

                    htmlWriter.NewLine = string.Empty;

                    htmlWriter.AddAttribute(CanvasControlAttribute, "");
                    htmlWriter.AddAttribute(CanvasDataVersionAttribute, this.DataVersion);
                    htmlWriter.AddAttribute(ControlDataAttribute, jsonControlData);
                    htmlWriter.RenderBeginTag(HtmlTextWriterTag.Div);
                    htmlWriter.RenderEndTag();
                }
            }

            return html.ToString();
        }
        #endregion
    }
    #endregion

    #region Available web part collection retrieved via _api/web/GetClientSideWebParts REST call
    /// <summary>
    /// Class holding a collection of client side webparts (retrieved via the _api/web/GetClientSideWebParts REST call)
    /// </summary>
    public class AvailableClientSideComponents
    {
        public ClientSideComponent[] value { get; set; }
    }

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
    #endregion
#endif
}
