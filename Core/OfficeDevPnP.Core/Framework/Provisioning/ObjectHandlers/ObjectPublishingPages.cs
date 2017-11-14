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
using System.Collections.Generic;

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
                web.EnsureProperty(w => w.WebTemplate);
                _willExtract = web.WebTemplate.ToLower().Contains("publishing") || web.WebTemplate.ToLower().Contains("enterprisewiki");
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

                            AddWebPartsToPublishingPage(web, page, mgr, parser);
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

        private static void AddWebPartsToPublishingPage(Web web, PublishingPage page, Microsoft.SharePoint.Client.WebParts.LimitedWebPartManager mgr, TokenParser parser)
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
                web.Context.Load(webPartProperties);
                web.Context.ExecuteQuery();

                if (wp.IsListViewWebPart)
                {
                    AddListViewWebpart(web, wp, definition, webPartProperties, parser);
                }
            }
        }

        private static void AddListViewWebpart(
            Web web,
            PublishingPageWebPart wp,
            Microsoft.SharePoint.Client.WebParts.WebPartDefinition definition,
            PropertyValues webPartProperties,
            TokenParser parser)
        {
            string defaultViewDisplayName = parser.ParseString(wp.DefaultViewDisplayName);

            if (!String.IsNullOrEmpty(defaultViewDisplayName))
            {
                string listUrl = webPartProperties.FieldValues["ListUrl"].ToString();

                web.Context.Load(definition, d => d.Id); // Id of the hidden view which gets automatically created
                web.Context.ExecuteQuery();

                Guid viewId = definition.Id;
                List list = web.GetListByUrl(listUrl);

                Microsoft.SharePoint.Client.View viewCreatedFromWebpart = list.Views.GetById(viewId);
                web.Context.Load(viewCreatedFromWebpart);

                Microsoft.SharePoint.Client.View viewCreatedFromList = list.Views.GetByTitle(defaultViewDisplayName);
                web.Context.Load(
                    viewCreatedFromList,
                    v => v.ViewFields,
                    v => v.ListViewXml,
                    v => v.ViewQuery,
                    v => v.ViewData,
                    v => v.ViewJoins,
                    v => v.ViewProjectedFields,
                    v => v.Paged,
                    v => v.RowLimit);

                web.Context.ExecuteQuery();

                //need to copy the same View definition to the new View added by the Webpart manager
                viewCreatedFromWebpart.ViewQuery = viewCreatedFromList.ViewQuery;
                viewCreatedFromWebpart.ViewData = viewCreatedFromList.ViewData;
                viewCreatedFromWebpart.ViewJoins = viewCreatedFromList.ViewJoins;
                viewCreatedFromWebpart.ViewProjectedFields = viewCreatedFromList.ViewProjectedFields;
                viewCreatedFromWebpart.Paged = viewCreatedFromList.Paged;
                viewCreatedFromWebpart.RowLimit = viewCreatedFromList.RowLimit;

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
            }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template,
            ProvisioningTemplateCreationInformation creationInfo)
        {
            var rootFolder = web.RootFolder;
            web.Context.Load(rootFolder);
            web.Context.ExecuteQueryRetry();
            string welcomePage = rootFolder.WelcomePage.Replace("Pages/", "").Replace(".aspx", "");
            template.Publishing = new Publishing();

            ClientContext ctx = web.Context.GetSiteCollectionContext();
            Web rootWeb = ctx.Site.RootWeb;
            ctx.Load(rootWeb);
            ctx.ExecuteQueryRetry();

            bool isRootWeb = (rootWeb.Id == web.Id);

            // if web is configured to only allow user to use specific page layout, we get layouts from the site property bag.
            template = GetAvailablePageLayouts(template, web, isRootWeb);

            // Get publishing pages of the current web
            template = GetPublishingPages(template, ctx, web, welcomePage, isRootWeb);

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
            if (Decimal.TryParse(item.ToString(), out Decimal number))
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

                        pageLayout.Path = layout.Attribute("url").Value;

                        if (pageLayout.Path == defaultPageLayoutUrl)
                        {
                            pageLayout.IsDefault = true;
                        }
                        template.Publishing.PageLayouts.Add(pageLayout);

                        if (isRootWeb)
                        {
                            var spFile = web.GetFileByServerRelativeUrl(web.ServerRelativeUrl + "/" + pageLayout.Path);
                            var fileStream = spFile.OpenBinaryStream();
                            web.Context.Load(spFile);
                            web.Context.Load(spFile.ListItemAllFields);
                            web.Context.ExecuteQuery();

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
                if (item.DisplayName == "tmp")
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
                            ctx.Load(wpd);
                            ctx.ExecuteQueryRetry();
                            var wp = wpd.WebPart;
                            ctx.Load(wp);
                            ctx.ExecuteQueryRetry();
                            PublishingPageWebPart ppwp = new PublishingPageWebPart();
                            ppwp.Title = wp.Title;
                            ppwp.Zone = wp.ZoneIndex.ToString();
                            var webPartId = wpd.Id;
                            var webPartXML = wpm.ExportWebPart(webPartId);
                            web.Context.ExecuteQuery();
                            ppwp.Contents = webPartXML.Value.Trim(new[] { '\n', ' ' });
                            web.Context.ExecuteQuery();
                            page.WebParts.Add(ppwp);
                        }
                        catch { }
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
                    // Check if page layout is not already defined in the PageLayout section of the template
                    if (!TemplateContainPageLayouts(template, url))
                    {
                        OfficeDevPnP.Core.Framework.Provisioning.Model.PageLayout targetFile =
                        new OfficeDevPnP.Core.Framework.Provisioning.Model.PageLayout
                        {
                            IsDefault = false,
                            Path = url
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
    }
}
