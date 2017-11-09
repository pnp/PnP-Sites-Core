using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using System.Net;
using System.IO;

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
                _willExtract = false;
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

                        if (page.WelcomePage && web.RootFolder.WelcomePage.Contains(parsedFullFileName))
                        {
                            //set the welcome page to a Temp page to allow page deletion
                            web.RootFolder.WelcomePage = "home.aspx";
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
                        var parentWeb = web.ParentWeb;
                        context.Load(parentWeb);
                        context.ExecuteQuery();
                        var RootWebUrl = parentWeb.ServerRelativeUrl;
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
    }
}
