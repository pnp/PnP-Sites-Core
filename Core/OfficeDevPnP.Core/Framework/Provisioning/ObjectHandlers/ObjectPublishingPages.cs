using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using System.Net;
using System.IO;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Xml.Linq;
using System.Xml;
using System.Text;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    class ObjectPublishingPages : ObjectHandlerBase
    {
        public override string Name { get { return "PublishingPages"; } }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Publishing != null && template.Publishing.PublishingPages.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                var publishingActivated = Boolean.Parse(web.GetPropertyBagValueString("__PublishingFeatureActivated", "false"));
                _willExtract = publishingActivated;
            }
            return _willExtract.Value;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser,
            ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);

                var siteContext = web.Context.GetSiteCollectionContext();
                var rootWeb = siteContext.Site.RootWeb;
                siteContext.Load(rootWeb);
                siteContext.ExecuteQueryRetry();

                // Upload available page layouts
                foreach (PageLayout pageLayout in template.Publishing.PageLayouts)
                {
                    var fileName = pageLayout.Path.Split('/').LastOrDefault();
                    var container = template.Connector.GetContainer();
                    var stream = template.Connector.GetFileStream(fileName, container);

                    // Get Masterpage catalog fodler
                    var masterpageCatalog = context.Site.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                    context.Load(masterpageCatalog);
                    context.ExecuteQuery();

                    var folder = masterpageCatalog.RootFolder;

                    UploadFile(template, pageLayout, folder, stream);
                }

                foreach (PublishingPage page in template.Publishing.PublishingPages)
                {

                    string parsedFileName = parser.ParseString(page.FileName);
                    string parsedFullFileName = parser.ParseString(page.FullFileName);

                    Microsoft.SharePoint.Client.Publishing.PublishingPage existingPage =
                        web.GetPublishingPage(parsedFileName + ".aspx");

                    if (!web.IsPropertyAvailable("RootFolder"))
                    {
                        web.Context.Load(web.RootFolder);
                        web.Context.ExecuteQueryRetry();
                    }

                    if (existingPage != null && existingPage.ServerObjectIsNull.Value == false)
                    {
                        if (page.WelcomePage && web.RootFolder.WelcomePage.ToLower().Contains(parsedFullFileName.ToLower()))
                        {
                            //set the welcome page to a Temp page to allow page deletion
                            web.RootFolder.WelcomePage = "temp.aspx";
                            web.RootFolder.Update();
                            web.Update();
                            context.ExecuteQueryRetry();
                        }
                        existingPage.ListItem.DeleteObject();
                        context.ExecuteQuery();
                    }


                    web.AddPublishingPage(
                        parsedFileName,
                        page.Layout,
                        parser.ParseString(page.Title)
                        );
                    Microsoft.SharePoint.Client.Publishing.PublishingPage publishingPage =
                        web.GetPublishingPage(parsedFullFileName);
                    Microsoft.SharePoint.Client.File pageFile = publishingPage.ListItem.File;
                    pageFile.CheckOut();

                    if (page.Properties != null && page.Properties.Count > 0)
                    {
                        context.Load(pageFile, p => p.Name, p => p.CheckOutType);
                        context.ExecuteQueryRetry();
                        var parsedProperties = page.Properties.ToDictionary(p => p.Key, p => parser.ParseString(p.Value));
                        var RootWebUrl = rootWeb.ServerRelativeUrl;
                        parsedProperties["PublishingPageLayout"] = RootWebUrl + parsedProperties["PublishingPageLayout"];
                        pageFile.SetFileProperties(parsedProperties, false);
                    }

                    if (page.WebParts != null && page.WebParts.Count > 0)
                    {
                        Microsoft.SharePoint.Client.WebParts.LimitedWebPartManager mgr =
                            pageFile.GetLimitedWebPartManager(
                                Microsoft.SharePoint.Client.WebParts.PersonalizationScope.Shared);
                        context.Load(mgr);
                        context.ExecuteQueryRetry();

                        AddWebPartsToPublishingPage(web, page, mgr, parser, publishingPage.ListItem);
                    }

                    List pagesLibrary = publishingPage.ListItem.ParentList;
                    context.Load(pagesLibrary);
                    context.ExecuteQueryRetry();

                    ListItem pageItem = publishingPage.ListItem;
                    web.Context.Load(pageItem, p => p.File.CheckOutType);
                    web.Context.ExecuteQueryRetry();

                    if (pageItem.File.CheckOutType != CheckOutType.None)
                    {
                        pageItem.File.CheckIn(String.Empty, CheckinType.MajorCheckIn);
                    }

                    if (page.Publish && pagesLibrary.EnableMinorVersions)
                    {
                        pageItem.File.Publish(String.Empty);
                        if (pagesLibrary.EnableModeration)
                        {
                            pageItem.File.Approve(String.Empty);
                        }
                    }


                    if (page.WelcomePage)
                    {
                        SetWelcomePage(web, pageFile);
                    }

                    context.ExecuteQueryRetry();
                }

                // Set available page layouts
                var availablePageLayouts = template.Publishing.PageLayouts.Select(p => p.Path);
                if (availablePageLayouts.Any())
                {
                    web.SetAvailablePageLayouts(rootWeb, availablePageLayouts);
                }

                // Set default page layout, if any
                var defaultPageLayout = template.Publishing.PageLayouts.FirstOrDefault(p => p.IsDefault);
                if (defaultPageLayout != null)
                {
                    web.SetDefaultPageLayoutForSite(rootWeb, defaultPageLayout.Path);
                }
            }
            return parser;
        }

        private static void SetWelcomePage(Web web, Microsoft.SharePoint.Client.File pageFile)
        {
            if (!web.IsPropertyAvailable("RootFolder"))
            {
                web.Context.Load(web.RootFolder);
                web.Context.ExecuteQueryRetry();
            }

            if (!pageFile.IsPropertyAvailable("ServerRelativeUrl"))
            {
                web.Context.Load(pageFile, p => p.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }

            var rootFolderRelativeUrl = pageFile.ServerRelativeUrl.Substring(web.RootFolder.ServerRelativeUrl.Length);

            web.SetHomePage(rootFolderRelativeUrl);
        }

        private static void AddWebPartsToPublishingPage(Web web, PublishingPage page, Microsoft.SharePoint.Client.WebParts.LimitedWebPartManager mgr, TokenParser parser, ListItem pageItem)
        {
            ClientContext ctx = web.Context as ClientContext;
            foreach (var wp in page.WebParts)
            {
                string wpContentsTokenResolved = parser.ParseString(wp.Contents).Replace("<property name=\"JSLink\" type=\"string\">" + ctx.Site.ServerRelativeUrl, "<property name=\"JSLink\" type=\"string\">~sitecollection");
                Microsoft.SharePoint.Client.WebParts.WebPart webPart = mgr.ImportWebPart(wpContentsTokenResolved).WebPart;
                Microsoft.SharePoint.Client.WebParts.WebPartDefinition definition = mgr.AddWebPart(
                                                                                            webPart,
                                                                                            wp.Zone,
                                                                                            (int)wp.Order
                                                                                        );
                var webPartProperties = definition.WebPart.Properties;
                web.Context.Load(definition.WebPart);
                web.Context.Load(definition, d  => d.Id);
                web.Context.Load(webPartProperties);
                web.Context.ExecuteQuery();

                if (wp.IsListViewWebPart)
                {
                    AddListViewWebpart(web, wp, definition, webPartProperties, parser);
                }

                if (wp.Zone == "wpz")
                {
                    //wpz means webpart inserted in PublishingPageContent
                    pageItem.Context.Load(pageItem, i => i["PublishingPageContent"]);
                    pageItem.Context.ExecuteQueryRetry();
                    pageItem["PublishingPageContent"] = String.Concat(pageItem["PublishingPageContent"], GetEmbeddedWPString(definition.Id));
                    pageItem.Update();
                    pageItem.Context.ExecuteQueryRetry();
                }
            }
        }

        private static string GetEmbeddedWPString(Guid wpGuid)
        {
            // set the web part's ID as part of the ID-s of the div elements
            string wpForm = @"
            <div class=""ms-rtestate-read ms-rte-wpbox"">
 
            <div class=""ms-rtestate-notify ms-rtegenerate-notify ms-rtestate-read {0}"" id=""div_{0}"">
                                        </div>
 
 
            <div id=""vid_{0}"" style=""display:none"">
                                        </div>
 
                                    </div>
 
            ";
            return string.Format(wpForm, wpGuid);
        }

        private static void AddListViewWebpart(
            Web web,
            PublishingPageWebPart wp,
            Microsoft.SharePoint.Client.WebParts.WebPartDefinition definition,
            PropertyValues webPartProperties,
            TokenParser parser)
        {
            string defaultViewDisplayName = parser.ParseString(wp.DefaultViewDisplayName);

            string listUrl = webPartProperties.FieldValues["ListUrl"].ToString();
            web.Context.Load(definition, d => d.Id); // Id of the hidden view which gets automatically created
            web.Context.ExecuteQuery();

            Guid viewId = definition.Id;
            List list = web.GetListByUrl(listUrl);
            web.Context.Load(list);
            web.Context.Load(list.Views);
            web.Context.ExecuteQuery();

            Microsoft.SharePoint.Client.View viewCreatedFromWebpart = list.Views.GetById(viewId);
            web.Context.Load(viewCreatedFromWebpart);

            //get xml node
            var existingViews = list.Views;
            web.Context.Load(existingViews, vs => vs.Include(v => v.Title, v => v.Id));
            web.Context.ExecuteQueryRetry();
            var currentViewIndex = 21;
            Microsoft.SharePoint.Client.View viewCreatedFromList = CreateView(web, wp.ViewContent, existingViews, list, currentViewIndex);

            //set basic view with defautviewdisplayname 
            //Microsoft.SharePoint.Client.View viewCreatedFromList = list.Views.GetByTitle(defaultViewDisplayName);

            web.Context.Load(
                    viewCreatedFromList,
                    v => v.ViewFields,
                    v => v.ListViewXml,
                    v => v.ViewQuery,
                    v => v.ViewData,
                    v => v.ViewJoins,
                    v => v.ViewProjectedFields,
                    v => v.Paged,
                    v => v.DefaultView,
                    v => v.RowLimit,
                    v => v.ContentTypeId,
                    v => v.Scope,
                    v => v.MobileView,
                    v => v.MobileDefaultView,
                    v => v.Aggregations,
                    v => v.JSLink,
                    v => v.ListViewXml,
                     v => v.StyleId
                    );

            web.Context.ExecuteQuery();

            //need to copy the same View definition to the new View added by the Webpart manager
            viewCreatedFromWebpart.ViewQuery = viewCreatedFromList.ViewQuery;
            viewCreatedFromWebpart.ViewData = viewCreatedFromList.ViewData;
            viewCreatedFromWebpart.ViewJoins = viewCreatedFromList.ViewJoins;
            viewCreatedFromWebpart.ViewProjectedFields = viewCreatedFromList.ViewProjectedFields;
            viewCreatedFromWebpart.Paged = viewCreatedFromList.Paged;
            viewCreatedFromWebpart.DefaultView = viewCreatedFromList.DefaultView;
            viewCreatedFromWebpart.RowLimit = viewCreatedFromList.RowLimit;
            viewCreatedFromWebpart.ContentTypeId = viewCreatedFromList.ContentTypeId;
            viewCreatedFromWebpart.Scope = viewCreatedFromList.Scope;
            viewCreatedFromWebpart.MobileView = viewCreatedFromList.MobileView;
            viewCreatedFromWebpart.MobileDefaultView = viewCreatedFromList.MobileDefaultView;
            viewCreatedFromWebpart.Aggregations = viewCreatedFromList.Aggregations;
            viewCreatedFromWebpart.JSLink = viewCreatedFromList.JSLink;
            viewCreatedFromWebpart.ListViewXml = viewCreatedFromList.ListViewXml;

            viewCreatedFromWebpart.ViewFields.RemoveAll();
            foreach (var field in viewCreatedFromList.ViewFields)
            {
                viewCreatedFromWebpart.ViewFields.Add(field);
            }

            if (webPartProperties.FieldValues.ContainsKey("JSLink") && webPartProperties.FieldValues["JSLink"] != null)
            {
                viewCreatedFromWebpart.JSLink = webPartProperties.FieldValues["JSLink"].ToString();
            }

            viewCreatedFromWebpart.Update();

            web.Context.ExecuteQuery();

            // remove view created for webpart
            web.Context.Load(list);
            web.Context.Load(list.Views);
            web.Context.ExecuteQuery();
            Microsoft.SharePoint.Client.View LRMCurrentView = list.Views.GetByTitle("LRMCurrentView");
            LRMCurrentView.DeleteObject();
            // Execute the query to the server    
            web.Context.ExecuteQuery();
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template,
            ProvisioningTemplateCreationInformation creationInfo)
        {
            var rootFolder = web.RootFolder;
            web.Context.Load(web, w => w.Id);
            web.Context.Load(rootFolder);
            web.Context.ExecuteQueryRetry();
            string welcomePage = rootFolder.WelcomePage.Replace("Pages/", "").Replace(".aspx", "");
            template.Publishing = new Publishing();

            ClientContext ctx = web.Context.GetSiteCollectionContext();
            Web rootWeb = ctx.Site.RootWeb;
            ctx.Load(rootWeb, w => w.Id);
            ctx.ExecuteQueryRetry();

            bool isRootWeb = (rootWeb.Id == web.Id);

            // if web is configured to only allow user to use specific page layout, we get layouts from the site property bag.
            template = GetAvailablePageLayouts(template, web, isRootWeb);

            // Get publishing pages of the current web
            template = GetPublishingPages(template, web.Context as ClientContext, web, welcomePage, isRootWeb);

            return template;
        }

        private static File UploadFile(ProvisioningTemplate template, Model.PageLayout file, Microsoft.SharePoint.Client.Folder folder, Stream stream)
        {
            File targetFile = null;
            var fileName = file.Path.Split('/').LastOrDefault(); ;
            try
            {
                targetFile = folder.UploadFile(fileName, stream, true);
            }
            catch (Exception)
            {
                //The file name might contain encoded characters that prevent upload. Decode it and try again.
                fileName = WebUtility.UrlDecode(fileName);
                targetFile = folder.UploadFile(fileName, stream, true);
            }
            return targetFile;
        }

        private const string AVAILABLEPAGELAYOUTS = "__PageLayouts";
        private const string DEFAULTPAGELAYOUT = "__DefaultPageLayout";
        private const string PAGE_LAYOUT_CONTENT_TYPE_ID = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE811";
        private const string HTML_PAGE_LAYOUT_CONTENT_TYPE_ID = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE8110003D357F861E29844953D5CAA1D4D8A3B";
        private const string MASTER_PAGE_CONTENT_TYPE_ID = "0x010105";
        private const string HTML_MASTER_PAGE_CONTENT_TYPE_ID = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F0054AD6EF48B9F7B45A142F8173F171BD10003D357F861E29844953D5CAA1D4D8A3A";
        private const string ASP_NET_MASTER_PAGE_CONTENT_TYPE_ID = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F0054AD6EF48B9F7B45A142F8173F171BD1";

        #region Private methods
        /// <summary>
        /// Return string according to the type of object passed as parameter
        /// </summary>
        /// <param name="item">The object</param>
        /// <returns></returns>
        private string returnString(object item)
        {
            if (item == null) { return ""; }
            if (item is FieldUserValue)
            {
                return ((FieldUserValue)item).LookupValue;
            }
            if (item is ContentTypeId)
            {
                return ((ContentTypeId)item).StringValue;
            }
            if (item is FieldUserValue[])
            {
                var lstuser = "";
                foreach (var user in (FieldUserValue[])item)
                    lstuser += user.LookupValue + ",";
                char[] charsToTrim = { ',' };
                return lstuser.Trim(charsToTrim);
            }
            if (item is FieldLookupValue)
            {
                return ((FieldLookupValue)item).LookupId.ToString();
            }
            if (item is FieldLookupValue[])
            {
                var flv = (FieldLookupValue[])item;
                string multilookup = "";
                foreach (var lookupValue in flv)
                {
                    multilookup += lookupValue.LookupId.ToString();
                }
                char[] charsToTrim = { ',' };
                multilookup.Trim(charsToTrim);
                return multilookup;
            }
            if (item is TaxonomyFieldValueCollection)
            {
                var lsttaxo = "";
                foreach (var taxo in ((TaxonomyFieldValueCollection)item).ToList().Select(t => t.Label).ToArray())
                    lsttaxo += taxo + ";";
                return lsttaxo;
            }
            if (item is TaxonomyFieldValue)
            {
                return ((TaxonomyFieldValue)item).Label;
            }
            if (item is FieldUrlValue)
            {
                return ((FieldUrlValue)item).Url + "," + ((FieldUrlValue)item).Description;
            }
            if (item is DateTime)
            {
                DateTime date = (DateTime)item;
                return date.ToLocalTime().ToString();
            }
            if (item is String)
            {
                return item.ToString();
            }
            if (item is Boolean)
            {
                return item.ToString();
            }
            if (item is Decimal)
            {
                return item.ToString().Replace(",", ".");
            }
            return "";
        }

        /// <summary>
        /// Get the specific page layouts defined in the property bag ot the current web
        /// if the web is configured to allow user to use onely specific page layouts
        /// </summary>
        /// <param name="web">The current web</param>
        /// <returns>The provisionning template</returns>
        private ProvisioningTemplate GetAvailablePageLayouts(ProvisioningTemplate template, Web web, bool isRootWeb)
        {
            var defaultLayoutXml = web.GetPropertyBagValueString(DEFAULTPAGELAYOUT, null);

            var defaultPageLayoutUrl = string.Empty;
            if (defaultLayoutXml != null && defaultLayoutXml.ToLower() != "__inherit")
            {
                defaultPageLayoutUrl = XElement.Parse(defaultLayoutXml).Attribute("url").Value.Replace("_catalogs/masterpage/", String.Empty);
            }

            var layoutsXml = web.GetPropertyBagValueString(AVAILABLEPAGELAYOUTS, null);

            if (!string.IsNullOrEmpty(layoutsXml) && layoutsXml.ToLower() != "__inherit")
            {
                var layoutsElement = XElement.Parse(layoutsXml);

                foreach (var layout in layoutsElement.Descendants("layout"))
                {
                    if (layout.Attribute("url") != null)
                    {
                        var pageLayout = new PageLayout();

                        var pageLayoutFullPath = layout.Attribute("url").Value;
                        pageLayout.Path = pageLayoutFullPath.Replace("_catalogs/masterpage/", String.Empty);

                        if (pageLayout.Path == defaultPageLayoutUrl)
                        {
                            pageLayout.IsDefault = true;
                        }
                        template.Publishing.PageLayouts.Add(pageLayout);

                        // Page layouts are always uploaded on root web
                        var siteContext = web.Context.GetSiteCollectionContext();
                        var rootWeb = siteContext.Site.RootWeb;
                        siteContext.Load(rootWeb);
                        siteContext.ExecuteQueryRetry();
                        rootWeb.EnsureProperty(w => w.ServerRelativeUrl);
                        var spFile = rootWeb.GetFileByServerRelativeUrl(rootWeb.ServerRelativeUrl + "/" + pageLayoutFullPath);
                        var fileStream = spFile.OpenBinaryStream();
                        rootWeb.Context.Load(spFile);
                        rootWeb.Context.Load(spFile.ListItemAllFields);
                        rootWeb.Context.ExecuteQuery();

                        try
                        {
                            template.Connector.SaveFileStream(spFile.Name, fileStream.Value);
                        }
                        catch (Exception)
                        {
                            //The file name might contain encoded characters that prevent upload. Decode it and try again.
                            var fileName = spFile.Name.Replace("&", "");
                            template.Connector.SaveFileStream(spFile.Name, fileStream.Value);
                        }
                    }

                }
            }
            return template;
        }

        /// <summary>
        /// Get publishing pages located in the Pages library of the current web
        /// </summary>
        /// <param name="template">Provisionning template</param>
        /// <param name="ctx">Client context</param>
        /// <param name="web">Current Web</param>
        /// <param name="welcomePage">Welcome page name</param>
        /// <returns>The provisionning template</returns>
        private ProvisioningTemplate GetPublishingPages(ProvisioningTemplate template, ClientContext ctx, Web web, string welcomePage, bool isRootWeb)
        {
            #region load needed context
            web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);
            string relativurl = web.ServerRelativeUrl;
            string url = web.Url;
            string pagesListId = web.GetPropertyBagValueString("__PagesListId", String.Empty);
            string namelayout = string.Empty;
            Web rootWeb = ctx.Site.RootWeb;
            ctx.Load(rootWeb);
            List pagesList = web.Lists.GetById(new Guid(pagesListId));
            ctx.Load(pagesList);
            ctx.ExecuteQueryRetry();
            ListItemCollection existingPages = pagesList.GetItems(CamlQuery.CreateAllItemsQuery());
            ctx.Load(existingPages, items => items.Include(item => item.DisplayName));
            ctx.ExecuteQueryRetry();
            var fieldColl = pagesList.Fields;
            ctx.Load(fieldColl);
            ctx.ExecuteQuery();
            #endregion

            foreach (ListItem item in existingPages)
            {
                ctx.Load(item);
                var readonlyFields = ctx.LoadQuery(pagesList.Fields.Where(f => f.Title == "Modified"));
                ctx.ExecuteQueryRetry();
                if (item.DisplayName == "tmp" || item.FileSystemObjectType == FileSystemObjectType.Folder)
                    continue;
                OfficeDevPnP.Core.Framework.Provisioning.Model.PublishingPage page = new OfficeDevPnP.Core.Framework.Provisioning.Model.PublishingPage { };

                #region Get Page properties
                Dictionary<string, string> properties = new Dictionary<string, string>();
                foreach (var property in item.FieldValues)
                {
                    var itemField = fieldColl.GetFieldByInternalName(property.Key);
                    if (itemField != null)
                    {
                        var val = returnString(property.Value);
                        if (!string.IsNullOrEmpty(val))
                        {
                            if (!itemField.ReadOnlyField)
                            {
                                val = val.Replace(url, "{site}");
                                val = val.Replace(relativurl, "{site}");
                                properties.Add(property.Key.ToString(), val);
                            }
                            if (property.Key == "PublishingPageLayout")
                            {
                                val = val.Replace(url, string.Empty).Replace(rootWeb.Url, string.Empty);
                                namelayout = val.Replace("/_catalogs/masterpage/", string.Empty).Split('.')[0];
                                properties.Add(property.Key.ToString(), val);
                                // Add associated page layout to the PageLayout section of the template
                                template = GetAssociatedPageLayouts(ctx, template, web, val, relativurl, isRootWeb);
                            }
                        }
                    }
                }
                page.Title = item.DisplayName;
                page.FileName = item.DisplayName;
                page.Layout = string.IsNullOrEmpty(namelayout) ? "BlankWebPartPage" : namelayout;
                page.Properties = properties;
                if (welcomePage.ToLower().Equals(item.DisplayName.ToLower()))
                    page.WelcomePage = true;
                page.PublishingPageContent = (string)item["PublishingPageContent"];
                #endregion

                #region Get Webparts
                var wpm = item.File.GetLimitedWebPartManager(Microsoft.SharePoint.Client.WebParts.PersonalizationScope.Shared);
                ctx.Load(wpm, webPartManager => webPartManager.WebParts, webPartManager => webPartManager.WebParts.Include(webPart => webPart.WebPart.Title));
                ctx.ExecuteQuery();
                if (wpm.WebParts != null && wpm.WebParts.AreItemsAvailable)
                {
                    foreach (var wpd in wpm.WebParts)
                    {
                        try
                        {
                            ctx.Load(wpd, def => def.ZoneId, def => def.Id);
                            ctx.ExecuteQueryRetry();
                            var wp = wpd.WebPart;
                            ctx.Load(wp, def => def.ZoneIndex, def => def.Title, def => def.Properties, def => def.TitleUrl);
                            ctx.ExecuteQueryRetry();
                            PublishingPageWebPart ppwp = new PublishingPageWebPart();
                            ppwp.Title = wp.Title;
                            ppwp.Order = (uint)wp.ZoneIndex;
                            ppwp.Zone = wpd.ZoneId.ToString();
                            var webPartId = wpd.Id;
                            if (!string.IsNullOrEmpty(wp.TitleUrl))
                            {
                                ppwp.DefaultViewDisplayName = "";
                                ppwp.ViewContent = wp.Properties.FieldValues["XmlDefinition"].ToString();

                                //set propertiess
                                string html = "";
                                Dictionary<string, string> propertiesWebpart = new Dictionary<string, string>();
                                foreach (var property in wp.Properties.FieldValues)
                                {
                                    var val = returnString(property.Value);
                                    var key = returnString(property.Key);
                                    if (!string.IsNullOrEmpty(val))
                                    {
                                        html += returnValue(key, val);
                                    }
                                }


                                ppwp.Contents = "<webParts>" +
                                " <webPart xmlns='http://schemas.microsoft.com/WebPart/v3'>" +
                                "    <metaData>" +
                                "      <type name='Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />" +
                                "      <importErrorMessage>Cannot import this Web Part.</importErrorMessage>" +
                                "    </metaData>" +
                                "    <data>" +
                                "      <properties>" +
                                "        <property name='ListUrl' type='string'>" + wp.TitleUrl.ToString().Replace(relativurl, "").Trim('/') + "</property>" +
                                            html +
                                "      </properties>" +
                                "    </data>" +
                                "  </webPart>" +
                                "</webParts>";

                            }
                            else {
                                // Webpart that are not checked as exportable would fail
                                wp.EnsureProperty(w => w.ExportMode);
                                if (wp.ExportMode != Microsoft.SharePoint.Client.WebParts.WebPartExportMode.All)
                                {
                                    bool forceCheckout = false;
                                    if (item.File.CheckOutType == CheckOutType.None)
                                    {
                                        item.File.CheckOut();
                                        ctx.ExecuteQuery();
                                        forceCheckout = true;
                                    }
                                    wp.ExportMode = Microsoft.SharePoint.Client.WebParts.WebPartExportMode.All;
                                    wpd.SaveWebPartChanges();

                                    if (forceCheckout)
                                    {
                                        item.File.CheckIn(String.Empty, CheckinType.MajorCheckIn);
                                        ctx.ExecuteQuery();
                                    }
                                }
                                var webPartXML = wpm.ExportWebPart(webPartId);
                                ctx.ExecuteQuery();
                                ppwp.ViewContent = "<View></View>";
                                ppwp.Contents = webPartXML.Value.Trim(new[] { '\n', ' ' }).Replace(web.Url, "{site}").Replace(web.ServerRelativeUrl, "{site}").Replace(rootWeb.Url, "{sitecollection}").Replace(rootWeb.ServerRelativeUrl,"{sitecollection}");
                                ctx.ExecuteQuery();
                            }
                            page.WebParts.Add(ppwp);
                        }
                        catch(Exception ex) { }
                    }
                }
                #endregion

                template.Publishing.PublishingPages.Add(page);
            }

            // Remove page file from the file node
            template.Files.RemoveAll(f => f.Folder.Contains("{site}/Pages"));

            return template;
        }

        /// <summary>
        /// Get the layout associated to a publishing page
        /// </summary>
        /// <param name="template">Provisionning template</param>
        /// <param name="web">Current web</param>
        /// <param name="val">Name of the associated publishing page</param>
        /// <returns></returns>
        private ProvisioningTemplate GetAssociatedPageLayouts(ClientContext context, ProvisioningTemplate template, Web web, string val, string relativeurl, bool isRootWeb)
        {
            try
            {
                string layoutName = val.Split(',')[1];
                string url = val.Split(',')[0];
                if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(layoutName))
                {
                    string layoutPath = url.Replace("_catalogs/masterpage/", String.Empty).TrimStart('/');
                    // Check if page layout is not already defined in the PageLayout section of the template
                    if (!TemplateContainPageLayouts(template, layoutPath))
                    {
                        OfficeDevPnP.Core.Framework.Provisioning.Model.PageLayout targetFile =
                        new OfficeDevPnP.Core.Framework.Provisioning.Model.PageLayout
                        {
                            IsDefault = false,
                            Path = layoutPath
                        };
                        template.Publishing.PageLayouts.Add(targetFile);

                        var rootWeb = web;
                        if (!isRootWeb)
                        {
                            rootWeb = context.Site.RootWeb;
                            context.Load(rootWeb);
                            context.Load(rootWeb, w => w.ServerRelativeUrl);
                            context.ExecuteQueryRetry();
                        }

                        var fileServerRelativeUrl = rootWeb.ServerRelativeUrl + url.Replace(web.Url, "").Replace(rootWeb.Url, "").Replace(web.ServerRelativeUrl, "").Replace(rootWeb.ServerRelativeUrl, "");
                        var spFile = rootWeb.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                        var fileStream = spFile.OpenBinaryStream();
                        context.Load(spFile);
                        context.Load(spFile.ListItemAllFields);
                        context.ExecuteQuery();
                        try
                        {
                            template.Connector.SaveFileStream(spFile.Name, fileStream.Value);
                        }
                        catch (Exception)
                        {
                            //The file name might contain encoded characters that prevent upload. Decode it and try again.
                            var fileName = spFile.Name.Replace("&", "");
                            template.Connector.SaveFileStream(spFile.Name, fileStream.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return template;
        }

        /// <summary>
        /// Check if page layout already exist in the PageLayouts section of the template
        /// </summary>
        /// <param name="template">Provisionning template</param>
        /// <param name="val">Nale of the layout</param>
        /// <returns></returns>
        private bool TemplateContainPageLayouts(ProvisioningTemplate template, string val)
        {
            foreach (var layout in template.Publishing.PageLayouts)
            {
                if (layout.Path == val)
                {
                    return true;
                }
            }
            return false;
        }
        #endregion
        private static Microsoft.SharePoint.Client.View CreateView(Web web, string view, Microsoft.SharePoint.Client.ViewCollection existingViews, List createdList, int currentViewIndex)
        {
            try
            {
                XElement viewElement = XElement.Parse(view);
                var displayNameElement = viewElement.Attribute("DisplayName");
                if (displayNameElement == null)
                {
                    throw new ApplicationException("Invalid View element, missing a valid value for the attribute DisplayName.");
                }

                var viewTitle = "LRMCurrentView";
                var existingView = existingViews.FirstOrDefault(v => v.Title == viewTitle);
                if (existingView != null)
                {
                    existingView.DeleteObject();
                    web.Context.ExecuteQueryRetry();
                }

                // Type
                var viewTypeString = viewElement.Attribute("Type") != null ? viewElement.Attribute("Type").Value : "None";
                viewTypeString = viewTypeString[0].ToString().ToUpper() + viewTypeString.Substring(1).ToLower();
                var viewType = (ViewType)Enum.Parse(typeof(ViewType), viewTypeString);

                // Fields
                string[] viewFields = null;
                var viewFieldsElement = viewElement.Descendants("ViewFields").FirstOrDefault();
                if (viewFieldsElement != null)
                {
                    viewFields = (from field in viewElement.Descendants("ViewFields").Descendants("FieldRef") select field.Attribute("Name").Value).ToArray();
                }

                // Default view
                var viewDefault = viewElement.Attribute("DefaultView") != null && bool.Parse(viewElement.Attribute("DefaultView").Value);

                // Row limit
                var viewPaged = true;
                uint viewRowLimit = 30;
                var rowLimitElement = viewElement.Descendants("RowLimit").FirstOrDefault();
                if (rowLimitElement != null)
                {
                    if (rowLimitElement.Attribute("Paged") != null)
                    {
                        viewPaged = bool.Parse(rowLimitElement.Attribute("Paged").Value);
                    }
                    viewRowLimit = uint.Parse(rowLimitElement.Value);
                }

                // Query
                var viewQuery = new StringBuilder();
                foreach (var queryElement in viewElement.Descendants("Query").Elements())
                {
                    viewQuery.Append(queryElement.ToString());
                }

                var viewCI = new ViewCreationInformation
                {
                    ViewFields = viewFields,
                    RowLimit = viewRowLimit,
                    Paged = viewPaged,
                    Title = viewTitle,
                    Query = viewQuery.ToString(),
                    ViewTypeKind = viewType,
                    PersonalView = false,
                    SetAsDefaultView = viewDefault,
                };

                // Allow to specify a custom view url. View url is taken from title, so we first set title to the view url value we need, 
                // create the view and then set title back to the original value
                var urlAttribute = viewElement.Attribute("Url");
                var urlHasValue = urlAttribute != null && !string.IsNullOrEmpty(urlAttribute.Value);
                if (urlHasValue)
                {
                    //set Title to be equal to url (in order to generate desired url)
                    viewCI.Title = Path.GetFileNameWithoutExtension(urlAttribute.Value);
                }

                var createdView = createdList.Views.Add(viewCI);
                createdView.EnsureProperties(v => v.Scope, v => v.JSLink, v => v.Title, v => v.Aggregations, v => v.MobileView, v => v.MobileDefaultView);
                web.Context.ExecuteQueryRetry();

                if (urlHasValue)
                {
                    //restore original title 
                    createdView.Title = viewTitle;
                    createdView.Update();
                }

                // ContentTypeID
                var contentTypeID = viewElement.Attribute("ContentTypeID") != null ? viewElement.Attribute("ContentTypeID").Value : null;
                if (!string.IsNullOrEmpty(contentTypeID) && (contentTypeID != BuiltInContentTypeId.System))
                {
                    ContentTypeId childContentTypeId = null;
                    if (contentTypeID == BuiltInContentTypeId.RootOfList)
                    {
                        var childContentType = web.GetContentTypeById(contentTypeID);
                        childContentTypeId = childContentType != null ? childContentType.Id : null;
                    }
                    else
                    {
                        childContentTypeId = createdList.ContentTypes.BestMatch(contentTypeID);
                    }
                    if (childContentTypeId != null)
                    {
                        createdView.ContentTypeId = childContentTypeId;
                        createdView.Update();
                    }
                }

                // Default for content type
                bool parsedDefaultViewForContentType;
                var defaultViewForContentType = viewElement.Attribute("DefaultViewForContentType") != null ? viewElement.Attribute("DefaultViewForContentType").Value : null;
                if (!string.IsNullOrEmpty(defaultViewForContentType) && bool.TryParse(defaultViewForContentType, out parsedDefaultViewForContentType))
                {
                    createdView.DefaultViewForContentType = parsedDefaultViewForContentType;
                    createdView.Update();
                }

                // Scope
                var scope = viewElement.Attribute("Scope") != null ? viewElement.Attribute("Scope").Value : null;
                ViewScope parsedScope = ViewScope.DefaultValue;
                if (!string.IsNullOrEmpty(scope) && Enum.TryParse<ViewScope>(scope, out parsedScope))
                {
                    createdView.Scope = parsedScope;
                    createdView.Update();
                }

                // MobileView
                var mobileView = viewElement.Attribute("MobileView") != null && bool.Parse(viewElement.Attribute("MobileView").Value);
                if (mobileView)
                {
                    createdView.MobileView = mobileView;
                    createdView.Update();
                }

                // MobileDefaultView
                var mobileDefaultView = viewElement.Attribute("MobileDefaultView") != null && bool.Parse(viewElement.Attribute("MobileDefaultView").Value);
                if (mobileDefaultView)
                {
                    createdView.MobileDefaultView = mobileDefaultView;
                    createdView.Update();
                }

                // Aggregations
                var aggregationsElement = viewElement.Descendants("Aggregations").FirstOrDefault();
                if (aggregationsElement != null)
                {
                    if (aggregationsElement.HasElements)
                    {
                        var fieldRefString = "";
                        var fieldRefs = aggregationsElement.Descendants("FieldRef");
                        foreach (var fieldRef in fieldRefs)
                        {
                            fieldRefString += fieldRef.ToString();
                        }
                        if (createdView.Aggregations != fieldRefString)
                        {
                            createdView.Aggregations = fieldRefString;
                            createdView.Update();
                        }
                    }
                }

                // ViewStyle
                var viewstyle = viewElement.Descendants("ViewStyle").FirstOrDefault();
                if (viewstyle != null)
                {
                    var viewStyleID = viewstyle.Attribute("ID") != null ? viewstyle.Attribute("ID").Value : "";
                    if (viewStyleID != "")
                    {
                        //parse xml
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml("<View><ViewStyle ID='" + viewStyleID.ToString() + "'/></View>");
                        XmlElement element = (XmlElement)doc.SelectSingleNode("//View//ViewStyle");
                        createdView.ListViewXml = doc.FirstChild.InnerXml;
                        createdView.Update();
                    }
                }

                // JSLink
                var jslinkElement = viewElement.Descendants("JSLink").FirstOrDefault();
                if (jslinkElement != null)
                {
                    var jslink = jslinkElement.Value;
                    if (createdView.JSLink != jslink)
                    {
                        createdView.JSLink = jslink;
                        createdView.Update();

                        // Only push the JSLink value to the web part as it contains a / indicating it's a custom one. So we're not pushing the OOB ones like clienttemplates.js or hierarchytaskslist.js
                        // but do push custom ones down to th web part (e.g. ~sitecollection/Style Library/JSLink-Samples/ConfidentialDocuments.js)
                        if (jslink.Contains("/"))
                        {
                            createdView.EnsureProperty(v => v.ServerRelativeUrl);
                            createdList.SetJSLinkCustomizations(createdView.ServerRelativeUrl, jslink);
                        }
                    }
                }


                createdList.Update();
                web.Context.ExecuteQueryRetry();

                // Add ListViewId token parser
                createdView.EnsureProperty(v => v.Id);
                return createdView;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private string returnValue(string key, string val)
        {
            string htmlvalueType = "";
            if (key == "Title" || key == "Default" || key == "NoDefaultStyle" || key == "ViewContentTypeId" || key == "XmlDefinitionLink" || key == "Description" || key == "JSLink" || key == "CatalogIconImageUrl" || key == "TitleIconImageUrl" || key == "Width" || key == "Height" || key == "HelpUrl" || key == "SelectParameters")
            {
                htmlvalueType = "<property name='" + key + "' type='string'>" + val + "</property>";
            }
            else if (key == "ChromeType")
            {
                string chrometype = Enum.GetName(typeof(System.Web.UI.WebControls.WebParts.PartChromeType), int.Parse(val));
                htmlvalueType = "<property name='" + key + "' type='chrometype'>" + chrometype + "</property>";
            }
            else if (key == "ShowWithSampleData" || key == "CacheXslStorage" || key == "ManualRefresh" || key == "EnableOriginalValue" || key == "ServerRender" || key == "AllowConnect" || key == "AllowZoneChange" || key == "DisableSaveAsNewViewButton" || key == "AutoRefresh" || key == "FireInitialRow" || key == "AllowEdit" || key == "AllowMinimize" || key == "UseSQLDataSourcePaging" || key == "ShowTimelineIfAvailable" || key == "Hidden" || key == "AllowClose" || key == "InplaceSearchEnabled" || key == "DisableViewSelectorMenu" || key == "IsClientRender" || key == "AsyncRefresh" || key == "HasClientDataSource")
            {
                htmlvalueType = "<property name='" + key + "' type='bool'>" + val + "</property>";
            }
            return htmlvalueType;
        }
    }
}
