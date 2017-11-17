using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using System.Net;
using System.IO;
using System.Xml;
using System.Text;
using System.Xml.Linq;

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
                    if (viewStyleID != "") { 
                        //parse xml
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml("<View><ViewStyle ID='"+viewStyleID.ToString()+"'/></View>");
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
    }
}
