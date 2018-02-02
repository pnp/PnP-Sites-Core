using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Utilities;
using System.Xml.Linq;
using System.Xml;
using System.Text;
using System.IO;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPages : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Pages"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.RootFolder.WelcomePage);

                // Check if this is not a noscript site as we're not allowed to update some properties
                bool isNoScriptSite = web.IsNoScriptSite();

                foreach (var page in template.Pages)
                {
                    var url = parser.ParseString(page.Url);

                    if (!url.ToLower().StartsWith(web.ServerRelativeUrl.ToLower()))
                    {
                        url = UrlUtility.Combine(web.ServerRelativeUrl, url);
                    }

                    var exists = true;
                    Microsoft.SharePoint.Client.File file = null;
                    try
                    {
                        file = web.GetFileByServerRelativeUrl(url);
                        web.Context.Load(file);
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (ServerException ex)
                    {
                        if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                        {
                            exists = false;
                        }
                    }
                    if (exists)
                    {
                        if (page.Overwrite)
                        {
                            try
                            {
                                scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0_, url);

                                // determine url of current home page
                                string welcomePageUrl = web.RootFolder.WelcomePage;
                                string welcomePageServerRelativeUrl = welcomePageUrl != null
                                    ? UrlUtility.Combine(web.ServerRelativeUrl, web.RootFolder.WelcomePage)
                                    : null;

                                bool overwriteWelcomePage = string.Equals(url, welcomePageServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase);

                                // temporarily reset home page so we can delete it
                                if (overwriteWelcomePage)
                                {
                                    web.SetHomePage(string.Empty);
                                }

                                file.DeleteObject();
                                web.Context.ExecuteQueryRetry();
                                web.AddWikiPageByUrl(url);
                                if (page.Layout == WikiPageLayout.Custom)
                                {
                                    web.AddLayoutToWikiPage(WikiPageLayout.OneColumn, url);
                                }
                                else
                                {
                                    web.AddLayoutToWikiPage(page.Layout, url);
                                }

                                if (overwriteWelcomePage)
                                {
                                    // restore welcome page to previous value
                                    web.SetHomePage(welcomePageUrl);
                                }
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Overwriting_existing_page__0__failed___1_____2_, url, ex.Message, ex.StackTrace);
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Pages_Creating_new_page__0_, url);

                            web.AddWikiPageByUrl(url);
                            if (page.Layout == WikiPageLayout.Custom)
                            {
                                web.AddLayoutToWikiPage(WikiPageLayout.OneColumn, url);
                            }
                            else
                            {
                                web.AddLayoutToWikiPage(page.Layout, url);
                            }
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Pages_Creating_new_page__0__failed___1_____2_, url, ex.Message, ex.StackTrace);
                        }
                    }

#pragma warning disable 618
                    if (page.WelcomePage)
#pragma warning restore 618
                    {
                        web.RootFolder.EnsureProperty(p => p.ServerRelativeUrl);
                        var rootFolderRelativeUrl = url.Substring(web.RootFolder.ServerRelativeUrl.Length);
                        web.SetHomePage(rootFolderRelativeUrl);
                    }

#if !SP2013
                    bool webPartsNeedLocalization = false;
#endif
                    if (page.WebParts != null & page.WebParts.Any())
                    {
                        if (!isNoScriptSite)
                        {
                            var existingWebParts = web.GetWebParts(url);

                            foreach (var webPart in page.WebParts)
                            {
                                if (existingWebParts.FirstOrDefault(w => w.WebPart.Title == parser.ParseString(webPart.Title)) == null)
                                {
                                    WebPartEntity wpEntity = new WebPartEntity();
                                    wpEntity.WebPartTitle = parser.ParseString(webPart.Title);
                                    wpEntity.WebPartXml = parser.ParseString(webPart.Contents.Trim(new[] { '\n', ' ' }), "~sitecollection", "~site");
                                    //if (!wpEntity.WebPartXml.Contains("<webParts>") && !wpEntity.WebPartXml.Contains("<WebParts>"))
                                    //{
                                    //    if (wpEntity.WebPartXml.Contains("http://schemas.microsoft.com/WebPart/v2")) {
                                    //        wpEntity.WebPartXml = "<WebPart xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns='http://schemas.microsoft.com/WebPart/v2'>" + wpEntity.WebPartXml + "</WebPart>";
                                    //    }
                                    //    else {
                                    //        wpEntity.WebPartXml = "<webParts>" + wpEntity.WebPartXml + "</webParts>";
                                    //    }

                                    //}
                                    var wpd = web.AddWebPartToWikiPage(url, wpEntity, (int)webPart.Row, (int)webPart.Column, false);
#if !SP2013
                                    if (webPart.Title.ContainsResourceToken())
                                    {
                                        // update data based on where it was added - needed in order to localize wp title
#if !SP2016
                                        wpd.EnsureProperties(w => w.ZoneId, w => w.WebPart, w => w.WebPart.Properties);
                                        webPart.Zone = wpd.ZoneId;
                                        wpd.EnsureProperties(w => w.WebPart, w => w.WebPart.Properties);
#else
                                        wpd.EnsureProperties(w => w.WebPart, w => w.WebPart.Properties);
#endif
                                        webPart.Order = (uint)wpd.WebPart.ZoneIndex;
                                        webPartsNeedLocalization = true;
                                    }
                                    if (webPart.ViewContent != "<View xmlns=\"\"></View>")
                                    {
                                        wpd.EnsureProperties(w => w.ZoneId, w => w.WebPart, w => w.WebPart.Properties);
                                        wpd.EnsureProperties(w => w.WebPart, w => w.WebPart.Properties);
                                        var webPartProperties = wpd.WebPart.Properties;
                                        AddListViewWebpart(web, webPart, wpd, webPartProperties);
                                    }
#endif
                                }
                            }

                            // Remove any existing WebPartIdToken tokens in the parser that were added by other pages. They won't apply to this page,
                            // and they'll cause issues if this page contains web parts with the same name as web parts on other pages.
                            parser.Tokens.RemoveAll(t => t is WebPartIdToken);

                            var allWebParts = web.GetWebParts(url);
                            foreach (var webpart in allWebParts)
                            {
                                parser.AddToken(new WebPartIdToken(web, webpart.WebPart.Title, webpart.Id));
                            }
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_Pages_SkipAddingWebParts, page.Url);
                        }
                    }

#if !SP2013
                    if (webPartsNeedLocalization)
                    {
                        page.LocalizeWebParts(web, parser, scope);
                    }
#endif

                    file = web.GetFileByServerRelativeUrl(url);
                    file.EnsureProperty(f => f.ListItemAllFields);

                    if (page.Fields.Any())
                    {
                        var item = file.ListItemAllFields;
                        foreach (var fieldValue in page.Fields)
                        {
                            item[fieldValue.Key] = parser.ParseString(fieldValue.Value);
                        }
                        item.Update();
                        web.Context.ExecuteQueryRetry();
                    }
                    if (page.Security != null && page.Security.RoleAssignments.Count != 0)
                    {
                        web.Context.Load(file.ListItemAllFields);
                        web.Context.ExecuteQueryRetry();
                        file.ListItemAllFields.SetSecurity(parser, page.Security);
                    }
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Impossible to return all files in the site currently

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }


        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Pages.Any();
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

        private static void AddListViewWebpart(
       Web web,
       WebPart wp,
       Microsoft.SharePoint.Client.WebParts.WebPartDefinition definition,
       PropertyValues webPartProperties)
        {

            string listUrl = webPartProperties.FieldValues["TitleUrl"].ToString();
            web.Context.Load(definition, d => d.Id); // Id of the hidden view which gets automatically created
            web.Context.ExecuteQuery();

            Guid viewId = definition.Id;
            listUrl = listUrl.Replace(web.ServerRelativeUrl + "/", "");
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
            definition.SaveWebPartChanges();

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


    }
}
