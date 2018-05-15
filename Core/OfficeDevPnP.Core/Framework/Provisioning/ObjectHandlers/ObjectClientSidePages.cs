using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Linq;
using OfficeDevPnP.Core.Utilities.CanvasControl;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
#if !ONPREMISES
    internal class ObjectClientSidePages : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "ClientSidePages"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(w => w.ServerRelativeUrl);

                // determine pages library
                string pagesLibrary = "SitePages";

                List<string> preCreatedPages = new List<string>();

                var currentPageIndex = 0;
                // pre create the needed pages so we can fill the needed tokens which might be used later on when we put web parts on those pages
                foreach (var clientSidePage in template.ClientSidePages)
                {
                    string pageName = $"{System.IO.Path.GetFileNameWithoutExtension(clientSidePage.PageName)}.aspx";
                    string url = $"{pagesLibrary}/{pageName}";

                    // Write page level status messages, needed in case many pages are provisioned
                    currentPageIndex++;
                    WriteMessage($"ClientSidePage|Create {pageName}|{currentPageIndex}|{template.ClientSidePages.Count}", ProvisioningMessageType.Progress);

                    url = UrlUtility.Combine(web.ServerRelativeUrl, url);

                    var exists = true;
                    try
                    {
                        var file = web.GetFileByServerRelativeUrl(url);
                        web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativeUrl);
                        web.Context.ExecuteQueryRetry();

                        // Fill token
                        parser.AddToken(new PageUniqueIdToken(web, file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                        parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                    }
                    catch (ServerException ex)
                    {
                        if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                        {
                            exists = false;
                        }
                    }

                    if (!exists)
                    {
                        // Pre-create the page    
                        Pages.ClientSidePage page = web.AddClientSidePage(pageName);

                        // Set page layout now, because once it's set, it can't be changed.
                        if (!string.IsNullOrEmpty(clientSidePage.Layout))
                        {
                            if (clientSidePage.Layout.Equals("Article", StringComparison.InvariantCultureIgnoreCase))
                            {
                                page.LayoutType = Pages.ClientSidePageLayoutType.Article;
                            }
                            else if (clientSidePage.Layout.Equals("Home", StringComparison.InvariantCultureIgnoreCase))
                            {
                                page.LayoutType = Pages.ClientSidePageLayoutType.Home;
                            }
                        }

                        page.Save(pageName);

                        var file = web.GetFileByServerRelativeUrl(url);
                        web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativeUrl);
                        web.Context.ExecuteQueryRetry();

                        // Fill token
                        parser.AddToken(new PageUniqueIdToken(web, file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                        parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));

                        // Track that we pre-added this page
                        preCreatedPages.Add(url);
                    }
                }

                currentPageIndex = 0;
                // Iterate over the pages and create/update them
                foreach (var clientSidePage in template.ClientSidePages)
                {
                    string pageName = $"{System.IO.Path.GetFileNameWithoutExtension(clientSidePage.PageName)}.aspx";
                    string url = $"{pagesLibrary}/{pageName}";
                    // Write page level status messages, needed in case many pages are provisioned
                    currentPageIndex++;
                    WriteMessage($"ClientSidePage|{pageName}|{currentPageIndex}|{template.ClientSidePages.Count}", ProvisioningMessageType.Progress);

                    url = UrlUtility.Combine(web.ServerRelativeUrl, url);

                    var exists = true;
                    try
                    {
                        var file = web.GetFileByServerRelativeUrl(url);
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

                    Pages.ClientSidePage page = null;
                    if (exists)
                    {
                        if (clientSidePage.Overwrite || preCreatedPages.Contains(url))
                        {
                            // Get the existing page
                            page = web.LoadClientSidePage(pageName);
                            // Clear the page
                            page.ClearPage();
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePages_NoOverWrite, pageName);
                            continue;
                        }
                    }
                    else
                    {
                        // Create new client side page
                        page = web.AddClientSidePage(pageName);
                    }

                    // Set page title
                    if (page.PageTitle != clientSidePage.Title)
                    {
                        page.PageTitle = clientSidePage.Title;
                    }

                    // Page Header
                    if (clientSidePage.Header != null)
                    {
                        switch (clientSidePage.Header.Type)
                        {
                            case ClientSidePageHeaderType.None:
                                {
                                    page.RemovePageHeader();
                                    break;
                                }
                            case ClientSidePageHeaderType.Default:
                                {
                                    page.SetDefaultPageHeader();
                                    break;
                                }
                            case ClientSidePageHeaderType.Custom:
                                {
                                    var serverRelativeImageUrl = parser.ParseString(clientSidePage.Header.ServerRelativeImageUrl);
                                    if (clientSidePage.Header.TranslateX.HasValue && clientSidePage.Header.TranslateY.HasValue)
                                    {
                                        page.SetCustomPageHeader(serverRelativeImageUrl, clientSidePage.Header.TranslateX.Value, clientSidePage.Header.TranslateY.Value);
                                    }
                                    else
                                    {
                                        page.SetCustomPageHeader(serverRelativeImageUrl);
                                    }
                                    break;
                                }
                        }
                    }

                    // Set page layout
                    if (!string.IsNullOrEmpty(clientSidePage.Layout))
                    {
                        if (clientSidePage.Layout.Equals("Article", StringComparison.InvariantCultureIgnoreCase))
                        {
                            page.LayoutType = Pages.ClientSidePageLayoutType.Article;
                        }
                        else if (clientSidePage.Layout.Equals("Home", StringComparison.InvariantCultureIgnoreCase))
                        {
                            page.LayoutType = Pages.ClientSidePageLayoutType.Home;
                        }
                    }

                    // Load existing available controls
                    var componentsToAdd = page.AvailableClientSideComponents().ToList();

                    // if no section specified then add a default single column section
                    if (!clientSidePage.Sections.Any())
                    {
                        clientSidePage.Sections.Add(new CanvasSection() { Type = CanvasSectionType.OneColumn, Order = 10 });
                    }

                    int sectionCount = -1;
                    // Apply the "layout" and content
                    foreach (var section in clientSidePage.Sections)
                    {
                        sectionCount++;
                        switch (section.Type)
                        {
                            case CanvasSectionType.OneColumn:
                                page.AddSection(Pages.CanvasSectionTemplate.OneColumn, section.Order);
                                break;
                            case CanvasSectionType.OneColumnFullWidth:
                                page.AddSection(Pages.CanvasSectionTemplate.OneColumnFullWidth, section.Order);
                                break;
                            case CanvasSectionType.TwoColumn:
                                page.AddSection(Pages.CanvasSectionTemplate.TwoColumn, section.Order);
                                break;
                            case CanvasSectionType.ThreeColumn:
                                page.AddSection(Pages.CanvasSectionTemplate.ThreeColumn, section.Order);
                                break;
                            case CanvasSectionType.TwoColumnLeft:
                                page.AddSection(Pages.CanvasSectionTemplate.TwoColumnLeft, section.Order);
                                break;
                            case CanvasSectionType.TwoColumnRight:
                                page.AddSection(Pages.CanvasSectionTemplate.TwoColumnRight, section.Order);
                                break;
                            default:
                                page.AddSection(Pages.CanvasSectionTemplate.OneColumn, section.Order);
                                break;
                        }

                        // Add controls to the section
                        if (section.Controls.Any())
                        {
                            // Safety measure: reset column order to 1 for columns marked with 0 or lower
                            foreach (var control in section.Controls.Where(p => p.Column <= 0).ToList())
                            {
                                control.Column = 1;
                            }

                            foreach (CanvasControl control in section.Controls)
                            {
                                Pages.ClientSideComponent baseControl = null;

                                // Is it a text control?
                                if (control.Type == WebPartType.Text)
                                {
                                    Pages.ClientSideText textControl = new Pages.ClientSideText();
                                    if (control.ControlProperties.Any())
                                    {
                                        var textProperty = control.ControlProperties.First();
                                        textControl.Text = parser.ParseString(textProperty.Value);
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(control.JsonControlData))
                                        {
                                            var json = JsonConvert.DeserializeObject<Dictionary<string, string>>(control.JsonControlData);

                                            if (json.Count > 0)
                                            {
                                                textControl.Text = parser.ParseString(json.First().Value);
                                            }
                                        }
                                    }
                                    // Reduce column number by 1 due 0 start indexing
                                    page.AddControl(textControl, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                                }
                                // It is a web part
                                else
                                {
                                    // apply token parsing on the web part properties
                                    control.JsonControlData = parser.ParseString(control.JsonControlData);

                                    // perform processing of web part properties (e.g. include listid property based list title property)
                                    var webPartPostProcessor = CanvasControlPostProcessorFactory.Resolve(control);
                                    webPartPostProcessor.Process(control, page);

                                    // Is a custom developed client side web part (3rd party)
                                    if (control.Type == WebPartType.Custom)
                                    {
                                        if (!string.IsNullOrEmpty(control.CustomWebPartName))
                                        {
                                            baseControl = componentsToAdd.FirstOrDefault(p => p.Name.Equals(control.CustomWebPartName, StringComparison.InvariantCultureIgnoreCase));
                                        }
                                        else if (control.ControlId != Guid.Empty)
                                        {
                                            baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals($"{{{control.ControlId.ToString()}}}", StringComparison.CurrentCultureIgnoreCase));

                                            if (baseControl == null)
                                            {
                                                baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals(control.ControlId.ToString(), StringComparison.InvariantCultureIgnoreCase));
                                            }
                                        }
                                    }
                                    // Is an OOB client side web part (1st party)
                                    else
                                    {
                                        string webPartName = "";
                                        switch (control.Type)
                                        {
                                            case WebPartType.Image:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.Image);
                                                break;
                                            case WebPartType.BingMap:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.BingMap);
                                                break;
                                            case WebPartType.ContentEmbed:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.ContentEmbed);
                                                break;
                                            case WebPartType.ContentRollup:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.ContentRollup);
                                                break;
                                            case WebPartType.DocumentEmbed:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.DocumentEmbed);
                                                break;
                                            case WebPartType.Events:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.Events);
                                                break;
                                            case WebPartType.GroupCalendar:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.GroupCalendar);
                                                break;
                                            case WebPartType.Hero:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.Hero);
                                                break;
                                            case WebPartType.ImageGallery:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.ImageGallery);
                                                break;
                                            case WebPartType.LinkPreview:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.LinkPreview);
                                                break;
                                            case WebPartType.List:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.List);
                                                break;
                                            case WebPartType.NewsFeed:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.NewsFeed);
                                                break;
                                            case WebPartType.NewsReel:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.NewsReel);
                                                break;
                                            case WebPartType.PageTitle:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.PageTitle);
                                                break;
                                            case WebPartType.People:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.People);
                                                break;
                                            case WebPartType.PowerBIReportEmbed:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.PowerBIReportEmbed);
                                                break;
                                            case WebPartType.QuickChart:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.QuickChart);
                                                break;
                                            case WebPartType.QuickLinks:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.QuickLinks);
                                                break;
                                            case WebPartType.SiteActivity:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.SiteActivity);
                                                break;
                                            case WebPartType.VideoEmbed:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.VideoEmbed);
                                                break;
                                            case WebPartType.YammerEmbed:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.YammerEmbed);
                                                break;
                                            case WebPartType.CustomMessageRegion:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.CustomMessageRegion);
                                                break;
                                            case WebPartType.Divider:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.Divider);
                                                break;
                                            case WebPartType.MicrosoftForms:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.MicrosoftForms);
                                                break;
                                            case WebPartType.Spacer:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.Spacer);
                                                break;
                                            case WebPartType.ClientWebPart:
                                                webPartName = Pages.ClientSidePage.ClientSideWebPartEnumToName(Pages.DefaultClientSideWebParts.ClientWebPart);
                                                break;
                                        }

                                        baseControl = componentsToAdd.FirstOrDefault(p => p.Name.Equals(webPartName, StringComparison.InvariantCultureIgnoreCase));
                                    }

                                    if (baseControl != null)
                                    {
                                        Pages.ClientSideWebPart myWebPart = new Pages.ClientSideWebPart(baseControl)
                                        {
                                            Order = control.Order
                                        };

                                        // Reduce column number by 1 due 0 start indexing
                                        page.AddControl(myWebPart, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                                        // set properties using json string
                                        if (!String.IsNullOrEmpty(control.JsonControlData))
                                        {
                                            myWebPart.PropertiesJson = control.JsonControlData;
                                        }

                                        // set using property collection
                                        if (control.ControlProperties.Any())
                                        {
                                            // grab the "default" properties so we can deduct their types, needed to correctly apply the set properties
                                            var controlManifest = JObject.Parse(baseControl.Manifest);
                                            JToken controlProperties = null;
                                            if (controlManifest != null)
                                            {
                                                controlProperties = controlManifest.SelectToken("preconfiguredEntries[0].properties");
                                            }

                                            foreach (var property in control.ControlProperties)
                                            {
                                                Type propertyType = typeof(string);

                                                if (controlProperties != null)
                                                {
                                                    var defaultProperty = controlProperties.SelectToken(property.Key, false);
                                                    if (defaultProperty != null)
                                                    {
                                                        propertyType = Type.GetType($"System.{defaultProperty.Type.ToString()}");

                                                        if (propertyType == null)
                                                        {
                                                            if (defaultProperty.Type.ToString().Equals("integer", StringComparison.InvariantCultureIgnoreCase))
                                                            {
                                                                propertyType = typeof(int);
                                                            }
                                                        }
                                                    }
                                                }

                                                myWebPart.Properties[property.Key] = JToken.FromObject(Convert.ChangeType(parser.ParseString(property.Value), propertyType));
                                            }
                                        }
                                    }
                                    else
                                    {
                                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePages_BaseControlNotFound, control.ControlId, control.CustomWebPartName);
                                    }
                                }
                            }
                        }
                    }

                    // Persist the page
                    page.Save(pageName);

                    // Set commenting, ignore on pages of the type Home
                    if (page.LayoutType != Pages.ClientSidePageLayoutType.Home)
                    {
                        // Make it a news page if requested
                        if (clientSidePage.PromoteAsNewsArticle)
                        {
                            page.PromoteAsNewsArticle();
                        }
                    }

                    if (clientSidePage.EnableComments)
                    {
                        page.EnableComments();
                    }
                    else
                    {
                        page.DisableComments();
                    }

                    // Publish page 
                    if (clientSidePage.Publish)
                    {
                        page.Publish();
                    }

                }
            }

            WriteMessage("Done processing Client Side Pages", ProvisioningMessageType.Completed);
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (new PnPMonitoredScope(this.Name))
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

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.ClientSidePages.Any();
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
    }
#endif
}
