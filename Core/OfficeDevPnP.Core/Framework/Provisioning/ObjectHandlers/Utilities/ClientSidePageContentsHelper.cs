using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities
{
#if !ONPREMISES
    /// <summary>
    /// Helper class holding public methods that used by the client side page object handler. The purpose is to be able to reuse these public methods in a extensibility provider
    /// </summary>
    public class ClientSidePageContentsHelper
    {
        /// <summary>
        /// Extracts a client side page
        /// </summary>
        /// <param name="web">Web to extract the page from</param>
        /// <param name="template">Current provisioning template that will hold the extracted page</param>
        /// <param name="creationInfo">ProvisioningTemplateCreationInformation passed into the provisioning engine</param>
        /// <param name="scope">Scope used for logging</param>
        /// <param name="pageUrl">Url of the page to extract</param>
        /// <param name="pageName">Name of the page to extract</param>
        /// <param name="isHomePage">Is this a home page?</param>
        public void ExtractClientSidePage(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string pageUrl, string pageName, bool isHomePage)
        {
            try
            {
                var pageToExtract = web.LoadClientSidePage(pageName);

                if (pageToExtract.Sections.Count == 0 && pageToExtract.Controls.Count == 0 && isHomePage)
                {
                    // This is default home page which was not customized...and as such there's no page definition stored in the list item. We don't need to extact this page.
                    scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_DefaultHomePage);
                }
                else
                {
                    // Create the page
                    var extractedPageInstance = new ClientSidePage()
                    {
                        PageName = pageName,
                        PromoteAsNewsArticle = false,
                        Overwrite = true,
                        Publish = true,
                        Layout = pageToExtract.LayoutType.ToString(),
                        EnableComments = !pageToExtract.CommentsDisabled,
                    };

                    if(pageToExtract.PageHeader != null)
                    {
                        var extractedHeader = new ClientSidePageHeader()
                        {
                            Type = (ClientSidePageHeaderType)Enum.Parse(typeof(Pages.ClientSidePageHeaderType),pageToExtract.PageHeader.Type.ToString()),
                            ServerRelativeImageUrl = pageToExtract.PageHeader.ImageServerRelativeUrl,
                            TranslateX = pageToExtract.PageHeader.TranslateX,
                            TranslateY = pageToExtract.PageHeader.TranslateY
                        };
                        extractedPageInstance.Header = extractedHeader;
                    }

                    // Add the sections
                    foreach (var section in pageToExtract.Sections)
                    {
                        // Set order
                        var sectionInstance = new CanvasSection()
                        {
                            Order = section.Order,
                        };

                        // Set section type
                        switch (section.Type)
                        {
                            case Pages.CanvasSectionTemplate.OneColumn:
                                sectionInstance.Type = CanvasSectionType.OneColumn;
                                break;
                            case Pages.CanvasSectionTemplate.TwoColumn:
                                sectionInstance.Type = CanvasSectionType.TwoColumn;
                                break;
                            case Pages.CanvasSectionTemplate.TwoColumnLeft:
                                sectionInstance.Type = CanvasSectionType.TwoColumnLeft;
                                break;
                            case Pages.CanvasSectionTemplate.TwoColumnRight:
                                sectionInstance.Type = CanvasSectionType.TwoColumnRight;
                                break;
                            case Pages.CanvasSectionTemplate.ThreeColumn:
                                sectionInstance.Type = CanvasSectionType.ThreeColumn;
                                break;
                            case Pages.CanvasSectionTemplate.OneColumnFullWidth:
                                sectionInstance.Type = CanvasSectionType.OneColumnFullWidth;
                                break;
                            default:
                                sectionInstance.Type = CanvasSectionType.OneColumn;
                                break;
                        }

                        // Add controls to section
                        foreach (var column in section.Columns)
                        {
                            foreach (var control in column.Controls)
                            {
                                // Create control 
                                CanvasControl controlInstance = new CanvasControl()
                                {
                                    Column = column.Order,
                                    ControlId = control.InstanceId,
                                    Order = control.Order,
                                };

                                // Set control type
                                if (control.Type == typeof(Pages.ClientSideText))
                                {
                                    controlInstance.Type = WebPartType.Text;

                                    // Set text content
                                    controlInstance.ControlProperties = new System.Collections.Generic.Dictionary<string, string>(1)
                                        {
                                            { "Text", TokenizeJsonTextData(web, (control as Pages.ClientSideText).Text) }
                                        };
                                }
                                else
                                {
                                    // set ControlId to webpart id 
                                    controlInstance.ControlId = Guid.Parse((control as Pages.ClientSideWebPart).WebPartId);
                                    var webPartType = Pages.ClientSidePage.NameToClientSideWebPartEnum((control as Pages.ClientSideWebPart).WebPartId);
                                    switch (webPartType)
                                    {
                                        case Pages.DefaultClientSideWebParts.ContentRollup:
                                            controlInstance.Type = WebPartType.ContentRollup;
                                            break;
                                        case Pages.DefaultClientSideWebParts.BingMap:
                                            controlInstance.Type = WebPartType.BingMap;
                                            break;
                                        case Pages.DefaultClientSideWebParts.ContentEmbed:
                                            controlInstance.Type = WebPartType.ContentEmbed;
                                            break;
                                        case Pages.DefaultClientSideWebParts.DocumentEmbed:
                                            controlInstance.Type = WebPartType.DocumentEmbed;
                                            break;
                                        case Pages.DefaultClientSideWebParts.Image:
                                            controlInstance.Type = WebPartType.Image;
                                            break;
                                        case Pages.DefaultClientSideWebParts.ImageGallery:
                                            controlInstance.Type = WebPartType.ImageGallery;
                                            break;
                                        case Pages.DefaultClientSideWebParts.LinkPreview:
                                            controlInstance.Type = WebPartType.LinkPreview;
                                            break;
                                        case Pages.DefaultClientSideWebParts.NewsFeed:
                                            controlInstance.Type = WebPartType.NewsFeed;
                                            break;
                                        case Pages.DefaultClientSideWebParts.NewsReel:
                                            controlInstance.Type = WebPartType.NewsReel;
                                            break;
                                        case Pages.DefaultClientSideWebParts.PowerBIReportEmbed:
                                            controlInstance.Type = WebPartType.PowerBIReportEmbed;
                                            break;
                                        case Pages.DefaultClientSideWebParts.QuickChart:
                                            controlInstance.Type = WebPartType.QuickChart;
                                            break;
                                        case Pages.DefaultClientSideWebParts.SiteActivity:
                                            controlInstance.Type = WebPartType.SiteActivity;
                                            break;
                                        case Pages.DefaultClientSideWebParts.VideoEmbed:
                                            controlInstance.Type = WebPartType.VideoEmbed;
                                            break;
                                        case Pages.DefaultClientSideWebParts.YammerEmbed:
                                            controlInstance.Type = WebPartType.YammerEmbed;
                                            break;
                                        case Pages.DefaultClientSideWebParts.Events:
                                            controlInstance.Type = WebPartType.Events;
                                            break;
                                        case Pages.DefaultClientSideWebParts.GroupCalendar:
                                            controlInstance.Type = WebPartType.GroupCalendar;
                                            break;
                                        case Pages.DefaultClientSideWebParts.Hero:
                                            controlInstance.Type = WebPartType.Hero;
                                            break;
                                        case Pages.DefaultClientSideWebParts.List:
                                            controlInstance.Type = WebPartType.List;
                                            break;
                                        case Pages.DefaultClientSideWebParts.PageTitle:
                                            controlInstance.Type = WebPartType.PageTitle;
                                            break;
                                        case Pages.DefaultClientSideWebParts.People:
                                            controlInstance.Type = WebPartType.People;
                                            break;
                                        case Pages.DefaultClientSideWebParts.QuickLinks:
                                            controlInstance.Type = WebPartType.QuickLinks;
                                            break;
                                        case Pages.DefaultClientSideWebParts.CustomMessageRegion:
                                            controlInstance.Type = WebPartType.CustomMessageRegion;
                                            break;
                                        case Pages.DefaultClientSideWebParts.Divider:
                                            controlInstance.Type = WebPartType.Divider;
                                            break;
                                        case Pages.DefaultClientSideWebParts.MicrosoftForms:
                                            controlInstance.Type = WebPartType.MicrosoftForms;
                                            break;
                                        case Pages.DefaultClientSideWebParts.Spacer:
                                            controlInstance.Type = WebPartType.Spacer;
                                            break;
                                        case Pages.DefaultClientSideWebParts.ClientWebPart:
                                            controlInstance.Type = WebPartType.ClientWebPart;
                                            break;
                                        case Pages.DefaultClientSideWebParts.ThirdParty:
                                            controlInstance.Type = WebPartType.Custom;
                                            break;
                                        default:
                                            controlInstance.Type = WebPartType.Custom;
                                            break;
                                    }

                                    // set the control properties
                                    if ((control as Pages.ClientSideWebPart).ServerProcessedContent != null)
                                    {
                                        // If we have serverProcessedContent then also export that one, it's important as some controls depend on this information to be present
                                        string serverProcessedContent = (control as Pages.ClientSideWebPart).ServerProcessedContent.ToString(Formatting.None);
                                        controlInstance.JsonControlData = "{ \"serverProcessedContent\": " + serverProcessedContent + ", \"properties\": " + (control as Pages.ClientSideWebPart).PropertiesJson + "}";
                                    }
                                    else
                                    {
                                        controlInstance.JsonControlData = (control as Pages.ClientSideWebPart).PropertiesJson;
                                    }

                                    // Tokenize the JsonControlData
                                    controlInstance.JsonControlData = TokenizeJsonControlData(web, controlInstance.JsonControlData);

                                    // Export relevant files if this flag is set
                                    if (creationInfo.PersistBrandingFiles)
                                    {
                                        List<Guid> fileGuids = new List<Guid>();
                                        Dictionary<string, string> exportedFiles = new Dictionary<string, string>();
                                        Dictionary<string, string> exportedPages = new Dictionary<string, string>();

                                        // grab all the guids in the already tokenized json and check try to get them as a file
                                        string guidPattern = "\"[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}\"";
                                        Regex regexClientIds = new Regex(guidPattern);
                                        if (regexClientIds.IsMatch(controlInstance.JsonControlData))
                                        {
                                            foreach (Match guidMatch in regexClientIds.Matches(controlInstance.JsonControlData))
                                            {
                                                Guid uniqueId;
                                                if (Guid.TryParse(guidMatch.Value.Trim("\"".ToCharArray()), out uniqueId))
                                                {
                                                    fileGuids.Add(uniqueId);
                                                }
                                            }
                                        }
                                        // grab all the encoded guids in the already tokenized json and check try to get them as a file
                                        guidPattern = "=[a-fA-F0-9]{8}%2D([a-fA-F0-9]{4}%2D){3}[a-fA-F0-9]{12}";
                                        regexClientIds = new Regex(guidPattern);
                                        if (regexClientIds.IsMatch(controlInstance.JsonControlData))
                                        {
                                            foreach (Match guidMatch in regexClientIds.Matches(controlInstance.JsonControlData))
                                            {
                                                Guid uniqueId;
                                                if (Guid.TryParse(guidMatch.Value.TrimStart("=".ToCharArray()), out uniqueId))
                                                {
                                                    fileGuids.Add(uniqueId);
                                                }
                                            }
                                        }

                                        // Iterate over the found guids to see if they're exportable files
                                        foreach (var uniqueId in fileGuids)
                                        {
                                            try
                                            {
                                                if (!exportedFiles.ContainsKey(uniqueId.ToString()))
                                                {
                                                    // Try to see if this is a file
                                                    var file = web.GetFileById(uniqueId);
                                                    web.Context.Load(file, f => f.Level, f => f.ServerRelativeUrl);
                                                    web.Context.ExecuteQueryRetry();

                                                    // If we got here it's a file, let's grab the file's path and name
                                                    var baseUri = new Uri(web.Url);
                                                    var fullUri = new Uri(baseUri, file.ServerRelativeUrl);
                                                    var folderPath = HttpUtility.UrlDecode(fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
                                                    var fileName = HttpUtility.UrlDecode(fullUri.Segments[fullUri.Segments.Count() - 1]);

                                                    // Don't export aspx files as some web parts refer to other client side pages --> pages have to be either exported as well or already exist in the target site
                                                    if (!fileName.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase))
                                                    {
                                                        var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

                                                        // Avoid duplicate file entries
                                                        var fileAlreadyExported = template.Files.Where(p => p.Folder.Equals(templateFolderPath, StringComparison.CurrentCultureIgnoreCase) &&
                                                                                                            p.Src.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefault();
                                                        if (fileAlreadyExported == null)
                                                        {
                                                            // Add a File to the template
                                                            template.Files.Add(new Model.File()
                                                            {
                                                                Folder = templateFolderPath,
                                                                Src = $"{templateFolderPath}/{fileName}",
                                                                Overwrite = true,
                                                                Level = (Model.FileLevel)Enum.Parse(typeof(Model.FileLevel), file.Level.ToString())
                                                            });

                                                            // Export the file
                                                            PersistFile(web, creationInfo, scope, folderPath, fileName);

                                                            // Keep track of the exported file path and it's UniqueId
                                                            exportedFiles.Add(uniqueId.ToString(), file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (!exportedPages.ContainsKey(uniqueId.ToString()))
                                                        {
                                                            exportedPages.Add(uniqueId.ToString(), file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()));
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_ErrorDuringFileExport, ex.Message);
                                            }
                                        }

                                        // Tokenize based on the found files, use a different token for encoded guids do we can later on replace by a new encoded guid
                                        foreach (var exportedFile in exportedFiles)
                                        {
                                            controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedFile.Key.Replace("-", "%2D"), $"{{fileuniqueidencoded:{exportedFile.Value}}}", RegexOptions.IgnoreCase);
                                            controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedFile.Key, $"{{fileuniqueid:{exportedFile.Value}}}", RegexOptions.IgnoreCase);
                                        }
                                        foreach(var exportedPage in exportedPages)
                                        {
                                            controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedPage.Key.Replace("-", "%2D"), $"{{pageuniqueidencoded:{exportedPage.Value}}}", RegexOptions.IgnoreCase);
                                            controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedPage.Key, $"{{pageuniqueid:{exportedPage.Value}}}", RegexOptions.IgnoreCase);
                                        }
                                    }
                                }

                                // add control to section
                                sectionInstance.Controls.Add(controlInstance);
                            }
                        }

                        extractedPageInstance.Sections.Add(sectionInstance);
                    }

                    // Renumber the sections...when editing modern homepages you can end up with section with order 0.5 or 0.75...let's ensure we render section as of 1
                    int sectionOrder = 1;
                    foreach (var sectionInstance in extractedPageInstance.Sections)
                    {
                        sectionInstance.Order = sectionOrder;
                        sectionOrder++;
                    }

                    // Add the page to the template
                    template.ClientSidePages.Add(extractedPageInstance);

                    // Set the homepage
                    if (isHomePage)
                    {
                        if (template.WebSettings == null)
                        {
                            template.WebSettings = new WebSettings();                            
                        }

                        if (pageUrl.StartsWith(web.ServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase))
                        {
                            template.WebSettings.WelcomePage = pageUrl.Replace(web.ServerRelativeUrl + "/", "");
                        }
                        else
                        {
                            template.WebSettings.WelcomePage = pageUrl;
                        }                        
                    }
                }
            }
            catch (ArgumentException ex)
            {
                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_NoValidPage, ex.Message);
            }
        }

        #region Helper methods
        private void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string folderPath, string fileName)
        {
            if (creationInfo.FileConnector != null)
            {
                var fileConnector = creationInfo.FileConnector;
                SharePointConnector connector = new SharePointConnector(web.Context, web.Url, "dummy");
                Uri u = new Uri(web.Url);

                if (u.PathAndQuery != "/")
                {
                    if (folderPath.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1)
                    {
                        folderPath = folderPath.Replace(u.PathAndQuery, "");
                    }
                }

                folderPath = HttpUtility.UrlDecode(folderPath);
                String container = HttpUtility.UrlDecode(folderPath).Trim('/').Replace("/", "\\");
                String persistenceFileName = HttpUtility.UrlDecode(fileName);

                if (fileConnector.Parameters.ContainsKey(FileConnectorBase.CONTAINER))
                {
                    container = string.Concat(fileConnector.GetContainer(), container);
                }

                using (Stream s = connector.GetFileStream(persistenceFileName, folderPath))
                {
                    if (s != null)
                    {
                        creationInfo.FileConnector.SaveFileStream(
                            persistenceFileName, container, s);
                    }
                }
            }
            else
            {
                scope.LogError("No connector present to persist homepage");
            }
        }

        private string TokenizeJsonControlData(Web web, string json)
        {
            var lists = web.Lists;
            var site = (web.Context as ClientContext).Site;
            web.Context.Load(site, s => s.Id, s => s.GroupId);
            web.Context.Load(web, w => w.ServerRelativeUrl, w => w.Id, w => w.Url);
            web.Context.Load(lists, ls => ls.Include(l => l.Id, l => l.Title, l => l.Views.Include(v=>v.Id, v => v.Title)));
            web.Context.ExecuteQueryRetry();

            // Tokenize list and list view id's as they can be used by client side web parts (like the list web part)
            foreach (var list in lists)
            {
                json = Regex.Replace(json, list.Id.ToString(), $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}", RegexOptions.IgnoreCase);
                foreach(var view in list.Views)
                {
                    json = Regex.Replace(json, view.Id.ToString(), $"{{viewid:{System.Security.SecurityElement.Escape(list.Title)},{System.Security.SecurityElement.Escape(view.Title)}}}", RegexOptions.IgnoreCase);
                }
            }

            // Some webparts might already contains the site URL using ~sitecollection token (i.e: CQWP) - shouldn't be needed for client side web parts, but just in case
            json = Regex.Replace(json, "\"~sitecollection/(.)*\"", "\"{site}\"", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, "'~sitecollection/(.)*'", "'{site}'", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, ">~sitecollection/(.)*<", ">{site}<", RegexOptions.IgnoreCase);

            // HostUrl token replacement
            var uri = new Uri(web.Url);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}:{uri.Port}", "{hosturl}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}", "{hosturl}", RegexOptions.IgnoreCase);

            // Site token replacement, also replace "encoded" guids
            json = Regex.Replace(json, site.Id.ToString(), "{sitecollectionid}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, site.Id.ToString().Replace("-", "%2D"), "{sitecollectionidencoded}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, web.Id.ToString(), "{siteid}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, web.Id.ToString().Replace("-", "%2D"), "{siteidencoded}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, "(\"" + web.ServerRelativeUrl + ")(?!&)", "\"{site}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, "'" + web.ServerRelativeUrl, "'{site}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, ">" + web.ServerRelativeUrl, ">{site}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, web.ServerRelativeUrl, "{site}", RegexOptions.IgnoreCase);

            // Connected Office 365 group tokenization
            if (site.GroupId != null && !site.GroupId.Equals(Guid.Empty))
            {
                json = Regex.Replace(json, site.GroupId.ToString(), "{sitecollectionconnectedoffice365groupid}", RegexOptions.IgnoreCase);
            }

            return json;
        }
        private string TokenizeJsonTextData(Web web, string json)
        {
            web.Context.Load(web, w => w.ServerRelativeUrl, w => w.Id, w => w.Url);
            web.Context.ExecuteQueryRetry();

            // Only replace links to content, other content is considered to be part of the "Text"
            json = Regex.Replace(json, "href=\"" + web.ServerRelativeUrl, "href=\"{site}", RegexOptions.IgnoreCase);

            return json;
        }
        #endregion
    }
#endif
}
