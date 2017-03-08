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
        ContentRollup,
        BingMap,
        ContentEmbed,
        DocumentEmbed,
        Image,
        ImageGallery,
        LinkPreview,
        NewsFeed,
        NewsReel,
        PowerBIReportEmbed,
        QuickChart,
        SiteActivity,
        VideoEmbed,
        YammerEmbed
    }

    /// <summary>
    /// Represents a modern client side page with all it's contents
    /// </summary>
    public class ClientSidePage
    {
        #region variables
        public const string CanvasField = "CanvasContent1";
        public const string SitePagesFeatureId = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec";

        private ClientContext context;
        private string pageName;
        private string pagesLibrary;
        private ListItem pageListItem;
        private string sitePagesServerRelativeUrl;
        private bool securityInitialized = false;
        private string accessToken;
        private System.Collections.Generic.List<CanvasZone> zones = new System.Collections.Generic.List<CanvasZone>(1);
        private System.Collections.Generic.List<CanvasControl> controls = new System.Collections.Generic.List<CanvasControl>(5);
        #endregion

        #region construction
        public ClientSidePage()
        {
            this.zones.Add(new CanvasZone(this, CanvasZoneTemplate.OneColumn, 0));
            this.pagesLibrary = "SitePages";
        }

        public ClientSidePage(ClientContext cc) : this()
        {
            if (cc == null)
            {
                throw new ArgumentNullException("Passed ClientContext object cannot be null");
            }
            this.context = cc;
        }
        #endregion

        #region Properties
        public System.Collections.Generic.List<CanvasZone> Zones
        {
            get
            {
                return this.zones;
            }
        }

        public System.Collections.Generic.List<CanvasControl> Controls
        {
            get
            {
                return this.controls;
            }
        }

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

        public ListItem PageListItem
        {
            get
            {
                return this.pageListItem;
            }
        }

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
        #endregion

        #region public methods
        public void AddZone(CanvasZone zone)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }
            this.zones.Add(zone);
        }

        public void AddZone(CanvasZone zone, int order)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }
            zone.Order = order;
            this.zones.Add(zone);
        }

        public void AddControl(CanvasControl control)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }

            // add to defaultzone and section
            control.zone = this.DefaultZone;
            control.section = this.DefaultZone.DefaultSection;

            this.controls.Add(control);
        }

        public void AddControl(CanvasControl control, int order)
        {
            if (control == null)
            {
                throw new ArgumentNullException("Passed control cannot be null");
            }

            // add to defaultzone and section
            control.zone = this.DefaultZone;
            control.section = this.DefaultZone.DefaultSection;
            control.Order = order;

            this.controls.Add(control);
        }

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

        public void Delete()
        {
            if (this.pageListItem == null)
            {
                throw new ArgumentException($"Page {this.pageName} was not loaded/saved to SharePoint and therefore can't be deleted");
            }

            pageListItem.DeleteObject();
            this.Context.ExecuteQueryRetry();
        }

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

            ClientSidePage page = new ClientSidePage(cc);
            page.pageName = pageName;

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
            var html = item[ClientSidePage.CanvasField].ToString();

            if (String.IsNullOrEmpty(html))
            {
                throw new ArgumentException($"Page {pageName} is not a \"modern\" client side page");
            }

            page.pageListItem = item;

            page.LoadFromHtml(html);
            return page;
        }

        public void Save(string pageName = null)
        {
            // Save page contents to SharePoint
            if (this.Context == null)
            {
                throw new Exception("No valid ClientContext object connected, can't save this page to SharePoint");
            }

            // Grab pages library reference
            List pagesLibrary = this.Context.Web.GetListByUrl(this.PagesLibrary, p => p.RootFolder);

            // Build up server relative page url
            if (String.IsNullOrEmpty(this.sitePagesServerRelativeUrl))
            {
                this.sitePagesServerRelativeUrl = pagesLibrary.RootFolder.ServerRelativeUrl;
            }

            if (!String.IsNullOrEmpty(pageName))
            {
                this.pageName = pageName;
            }

            string serverRelativePageName = $"{this.sitePagesServerRelativeUrl}/{this.pageName}";

            // ensure page exists
            var pageFile = this.Context.Web.GetFileByServerRelativeUrl(serverRelativePageName);
            this.Context.Web.Context.Load(pageFile, f => f.ListItemAllFields, f => f.Exists);
            this.Context.Web.Context.ExecuteQueryRetry();

            ListItem item;
            if (!pageFile.Exists)
            {
                // create page listitem
                item = pagesLibrary.RootFolder.Files.AddTemplateFile(serverRelativePageName, TemplateFileType.ClientSidePage).ListItemAllFields;
                // Fix page to be modern
                item["ContentTypeId"] = BuiltInContentTypeId.ModernArticlePage;
                item["Title"] = System.IO.Path.GetFileNameWithoutExtension(this.pageName);
                item["ClientSideApplicationId"] = ClientSidePage.SitePagesFeatureId;
                item["PageLayoutType"] = "Article";
                item["PromotedState"] = "0";
                item["BannerImageUrl"] = "/_layouts/15/images/sitepagethumbnail.png";
                item.Update();
                this.Context.Web.Context.Load(item);
                this.Context.Web.Context.ExecuteQueryRetry();
            }
            else
            {
                item = pageFile.ListItemAllFields;
            }

            // Persist to page field
            item[ClientSidePage.CanvasField] = this.ToHtml();
            item.Update();
            this.Context.ExecuteQueryRetry();

            this.pageListItem = item;
        }

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

        public ClientSideWebPart InstantiateDefaultWebPart(DefaultClientSideWebParts webPart)
        {
            var webPartName = this.ClientSideWebPartEnumToName(webPart);
            var webParts = this.AvailableClientSideComponents(webPartName);

            if (webParts.Count() == 1)
            {
                return new ClientSideWebPart(webParts.First());
            }

            return null;
        }

        public System.Collections.Generic.IEnumerable<ClientSideComponent> AvailableClientSideComponents()
        {
            return this.AvailableClientSideComponents(null);
        }

        public System.Collections.Generic.IEnumerable<ClientSideComponent> AvailableClientSideComponents(DefaultClientSideWebParts webPart)
        {
            return this.AvailableClientSideComponents(this.ClientSideWebPartEnumToName(webPart));
        }

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
        #endregion

        #region Internal and private methods
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

                // assume it's one section with one zone
                int controlOrder = 0;
                foreach (var clientSideControl in clientSideControls)
                {
                    var controlData = clientSideControl.GetAttribute(CanvasControl.ControlDataAttribute);
                    var controlType = CanvasControl.GetType(controlData);

                    if (controlType == typeof(ClientSideText))
                    {
                        var control = new ClientSideText();
                        control.Order = controlOrder;
                        control.FromHtml(clientSideControl);
                        this.AddControl(control);
                    }
                    else if (controlType == typeof(ClientSideWebPart))
                    {
                        var control = new ClientSideWebPart();
                        control.FromHtml(clientSideControl);
                        this.AddControl(control);
                    }

                    controlOrder++;
                }
            }
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

        private string ClientSideWebPartEnumToName(DefaultClientSideWebParts webPart)
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
                default: return "";
            }
        }
        #endregion
    }

    /// <summary>
    /// The type of canvas being used
    /// </summary>
    public enum CanvasZoneTemplate
    {
        OneColumn = 0
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
            Type = CanvasZoneTemplate.OneColumn;
            Order = 0;
            this.sections.Add(new CanvasSection(this));
        }

        internal CanvasZone(ClientSidePage page, CanvasZoneTemplate canvasSectionTemplate, int order)
        {
            if (page == null)
            {
                throw new ArgumentNullException("Passed page cannot be null");
            }

            this.page = page;
            Type = canvasSectionTemplate;
            Order = order;

            if (canvasSectionTemplate == CanvasZoneTemplate.OneColumn)
            {
                this.sections.Add(new CanvasSection(this));
            }
        }
        #endregion

        #region Properties
        public CanvasZoneTemplate Type { get; set; }

        public int Order { get; set; }

        public System.Collections.Generic.List<CanvasSection> Sections
        {
            get
            {
                return this.sections;
            }
        }

        public ClientSidePage Page
        {
            get
            {
                return this.page;
            }
        }

        public System.Collections.Generic.List<CanvasControl> Controls
        {
            get
            {
                return this.Page.Controls.Where(p => p.Zone == this).ToList<CanvasControl>();
            }
        }

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
        public void AddSection(CanvasSection section)
        {
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            this.sections.Add(section);
        }

        public void AddSection(CanvasSection section, int order)
        {
            if (section == null)
            {
                throw new ArgumentNullException("Passed section cannot be null");
            }

            section.Order = order;
            this.sections.Add(section);
        }

        public string ToHtml()
        {
            StringBuilder html = new StringBuilder(100);
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;

                if (Type == CanvasZoneTemplate.OneColumn)
                {
                    foreach (var zone in this.sections.OrderBy(z => z.Order))
                    {
                        htmlWriter.Write(zone.ToHtml());
                    }
                }
            }

            return html.ToString();
        }
        #endregion
    }

    /// <summary>
    /// Represents a section in a canvas zone
    /// </summary>
    public class CanvasSection
    {
        #region variables
        private int sectionFactor;
        private CanvasZone zone;
        #endregion

        #region construction
        internal CanvasSection(CanvasZone zone)
        {
            if (zone == null)
            {
                throw new ArgumentNullException("Passed zone cannot be null");
            }

            this.zone = zone;
            this.sectionFactor = 0;
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
        public int Order { get; set; }

        public CanvasZone Zone
        {
            get
            {
                return this.zone;
            }
        }

        public int SectionFactor
        {
            get
            {
                return this.sectionFactor;
            }
        }

        public System.Collections.Generic.List<CanvasControl> Controls
        {
            get
            {
                return this.Zone.Page.Controls.Where(p => p.Zone == this.Zone && p.Section == this).ToList<CanvasControl>();
            }
        }
        #endregion

        #region public methods
        public string ToHtml()
        {
            StringBuilder html = new StringBuilder(100);
            using (var htmlWriter = new HtmlTextWriter(new System.IO.StringWriter(html), ""))
            {
                htmlWriter.NewLine = string.Empty;

                foreach (var control in this.Zone.Page.Controls.Where(p => p.Zone == this.Zone && p.Section == this).OrderBy(z => z.Order))
                {
                    htmlWriter.Write(control.ToHtml());
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
        public int ComponentType { get; set; }
        public string Id { get; set; }
        public string Manifest { get; set; }
        public int ManifestType { get; set; }
        public string Name { get; set; }
        public int Status { get; set; }
    }
    #endregion
#endif
}
