using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

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

                foreach (PublishingPage page in template.Publishing.PublishingPages)
                {
                    Microsoft.SharePoint.Client.Publishing.PublishingPage existingPage =
                        web.GetPublishingPage(page.FileName + ".aspx");

                    if (!web.IsPropertyAvailable("RootFolder"))
                    {
                        web.Context.Load(web.RootFolder);
                        web.Context.ExecuteQueryRetry();
                    }

                    if (existingPage != null && existingPage.ServerObjectIsNull.Value == false)
                    {
                        if (!page.Overwrite)
                        {
                            scope.LogDebug(
                                CoreResources.Provisioning_ObjectHandlers_PublishingPages_Skipping_As_Overwrite_false,
                                page.FileName);
                            continue;
                        }

                        if (page.WelcomePage && web.RootFolder.WelcomePage.Contains(page.FullFileName))
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
                        page.FileName,
                        page.Layout,
                        parser.ParseString(page.Title)
                        );
                    Microsoft.SharePoint.Client.Publishing.PublishingPage publishingPage =
                        web.GetPublishingPage(page.FullFileName);
                    Microsoft.SharePoint.Client.File pageFile = publishingPage.ListItem.File;
                    pageFile.CheckOut();

                    if (page.Properties != null && page.Properties.Count > 0)
                    {
                        context.Load(pageFile, p => p.Name, p => p.CheckOutType);
                        context.ExecuteQueryRetry();
                        var parsedProperties = page.Properties.ToDictionary(p => p.Key, p => parser.ParseString(p.Value));
                        pageFile.SetFileProperties(parsedProperties, false);
                    }

                    if (page.WebParts != null && page.WebParts.Count > 0)
                    {
                        Microsoft.SharePoint.Client.WebParts.LimitedWebPartManager mgr =
                            pageFile.GetLimitedWebPartManager(
                                Microsoft.SharePoint.Client.WebParts.PersonalizationScope.Shared);
                        context.Load(mgr);
                        context.ExecuteQueryRetry();

                        AddWebPartsToPublishingPage(page, context, mgr);
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

        private static void AddWebPartsToPublishingPage(PublishingPage page, ClientContext ctx, Microsoft.SharePoint.Client.WebParts.LimitedWebPartManager mgr)
        {
            foreach (var wp in page.WebParts)
            {
                string wpContentsTokenResolved = wp.Contents;
                Microsoft.SharePoint.Client.WebParts.WebPart webPart = mgr.ImportWebPart(wpContentsTokenResolved).WebPart;
                Microsoft.SharePoint.Client.WebParts.WebPartDefinition definition = mgr.AddWebPart(
                                                                                            webPart,
                                                                                            wp.Zone,
                                                                                            (int)wp.Order
                                                                                        );
                var webPartProperties = definition.WebPart.Properties;
                ctx.Load(definition.WebPart);
                ctx.Load(webPartProperties);
                ctx.ExecuteQuery();

                if (wp.IsListViewWebPart)
                {
                    AddListViewWebpart(ctx, wp, definition, webPartProperties);
                }
            }
        }

        private static void AddListViewWebpart(
            ClientContext ctx,
            PublishingPageWebPart wp,
            Microsoft.SharePoint.Client.WebParts.WebPartDefinition definition,
            PropertyValues webPartProperties)
        {
            string defaultViewDisplayName = wp.DefaultViewDisplayName;

            if (!String.IsNullOrEmpty(defaultViewDisplayName))
            {
                string listUrl = webPartProperties.FieldValues["ListUrl"].ToString();

                ctx.Load(definition, d => d.Id); // Id of the hidden view which gets automatically created
                ctx.ExecuteQuery();

                Guid viewId = definition.Id;
                List list = ctx.Web.GetListByUrl(listUrl);

                Microsoft.SharePoint.Client.View viewCreatedFromWebpart = list.Views.GetById(viewId);
                ctx.Load(viewCreatedFromWebpart);

                Microsoft.SharePoint.Client.View viewCreatedFromList = list.Views.GetByTitle(defaultViewDisplayName);
                ctx.Load(
                    viewCreatedFromList,
                    v => v.ViewFields,
                    v => v.ListViewXml,
                    v => v.ViewQuery,
                    v => v.ViewData,
                    v => v.ViewJoins,
                    v => v.ViewProjectedFields);

                ctx.ExecuteQuery();

                //need to copy the same View definition to the new View added by the Webpart manager
                viewCreatedFromWebpart.ViewQuery = viewCreatedFromList.ViewQuery;
                viewCreatedFromWebpart.ViewData = viewCreatedFromList.ViewData;
                viewCreatedFromWebpart.ViewJoins = viewCreatedFromList.ViewJoins;
                viewCreatedFromWebpart.ViewProjectedFields = viewCreatedFromList.ViewProjectedFields;
                viewCreatedFromWebpart.ViewFields.RemoveAll();

                foreach (var field in viewCreatedFromList.ViewFields)
                {
                    viewCreatedFromWebpart.ViewFields.Add(field);
                }

                //need to set the JSLink to the new View added by the Webpart manager.
                //This is because there's no way to change the BaseViewID property of the new View,
                //and we needed to do that because the custom JSLink was bound to a specific BaseViewID (overrideCtx.BaseViewID = 3;)
                //The work around to this is to add the JSLink to the specific new View created when you add the xsltViewWebpart to the page
                //and remove the "overrideCtx.BaseViewID = 3;" from the JSLink file
                //that way, the JSLink will be executed only for this View, that is only used in the xsltViewWebpart,
                //so the effect is the same that bind the JSLink to the BaseViewID
                if (webPartProperties.FieldValues.ContainsKey("JSLink") && webPartProperties.FieldValues["JSLink"] != null)
                {
                    viewCreatedFromWebpart.JSLink = webPartProperties.FieldValues["JSLink"].ToString();
                }

                viewCreatedFromWebpart.Update();

                ctx.ExecuteQuery();
            }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template,
            ProvisioningTemplateCreationInformation creationInfo)
        {
            return template;
        }
    }
}
